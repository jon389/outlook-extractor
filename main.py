import pandas as pd, numpy as np, xlwings as xw
from pathlib import Path
import win32com.client
from win32com.client import constants
from datetime import datetime
from dateutil import tz
import hashlib
import dataset
from log_conf import logger as log

save_attachments_temp = Path(__file__).parent / 'attachments_temp'
if not save_attachments_temp.exists():
    save_attachments_temp.mkdir()
for f in save_attachments_temp.iterdir():  # delete all files
    f.unlink()

mock_emails_db = Path(__file__).parent / 'parsed_mock_emails.db'


def match_mock_email(From: str, To: str, Subject: str, attached_filenames: list[str]) -> list[str]:
    if ('jon' in From.lower() and
            'jon' in To.lower() and
            'outlook' in Subject.lower() and 'extract' in Subject.lower()):
        matched_attachments = [a for a in attached_filenames if ('mock' in a and 'attach' in a
                                                                 and '.xls' in a)]  # matches xls, xlsx, xlsm, xlsb
        return matched_attachments


def parse_mock_attachment_to_db(mailitem, attachment_name: str, attachment_temp_filename: str):
    # ParsedEmails Table
    # ParseTimestampUTC
    # mailitem: From Subject ReceivedTime Body
    # Attachments -> list of attachment file names
    # SelectedAttachment -> file name of selected parsed xlsx
    # SelectedAttachmentHash

    # check if already parsed SelectedAttachment/SelectedAttachmentCheckSum

    db = dataset.connect(f'sqlite:///{mock_emails_db}')
    parsed_emails = db['ParsedEmails']
    data_table = db['ParsedAttachmentData']

    expected_data_cols = dict(
        datetime='Date'.split(),
        meta='first_name last_name email gender ip_address'.split(),
        decimal='GST PST HST JST KST'.split(),
        other='Comment'.split(),
    )
    expected_data_cols_all = [col for col_grp in expected_data_cols.values() for col in col_grp]

    parse_timestamp = datetime.utcnow()  # datetime.now(tz.tzlocal())
    attachment_hash = hashlib.sha1(open(attachment_temp_filename, 'rb').read()).hexdigest()

    already_exists = next(parsed_emails.find(Attachment=attachment_name, AttachmentHash=attachment_hash), None)
    if already_exists:
        log.info(f'ignoring {attachment_name} from {mailitem.Sender.Name}|{mailitem.Subject:.40}|'
                    f'{mailitem.ReceivedTime:%Y-%m-%d %H:%M}'
                 f' already in ParsedEmails table '
                 f'from {already_exists["From"]}|{already_exists["Subject"]:.40}|'
                 f'{already_exists["Received"]:%Y-%m-%d %H:%M}')
        return

    db.begin()
    try:
        parsed_emails.insert(dict(
            ParseTimestampUTC=parse_timestamp,
            To=mailitem.To,
            From=mailitem.Sender.Name,
            Subject=mailitem.Subject,
            Received=mailitem.ReceivedTime,
            Size=mailitem.Size,
            Body=mailitem.Body,
            EmailAttachments=','.join(a.FileName for a in mailitem.Attachments),
            Attachment=attachment_name,
            AttachmentHash=attachment_hash
        ))

        attach_data = pd.read_excel(attachment_temp_filename,
                                    parse_dates=['Date'])
        # check data consistency
        #   eg. expected columns exist
        assert all(col in attach_data for col in expected_data_cols_all)
        #   eg. Date col is parsed as a date
        assert 'datetime' in str(attach_data['Date'].dtype)
        #   eg. that input data has unique (Date, email)
        if attach_data.duplicated('Date email'.split()).any():
            err_msg = f'data error, duplicate rows in {attachment_name}'
            log.error(err_msg)
            raise ValueError(err_msg)

        # use external key to map to row of ParsedEmails-> ParseTimestampUTC From Subject ReceivedTime Attachment
        # don't insert duplicate data rows (ignoring ParseTimestampUTC)
        #   Declare a unique constraint? on (<list of columns>).
        # data_cols = 'Date first_name last_name email gender ip_address GST PST HST JST KST Comment'
        for idx, row in attach_data.iterrows():
            already_exists = next(data_table.find(
                Date=row['Date'],
                email=row['email'],
                order_by=['-ParseTimestampUTC']
            ), None)

            has_data_revision = False
            if already_exists:
                log.debug(f'data row Date {row["Date"]:%Y-%m-%d} {row["email"]} already exist')
                if not (all(already_exists[meta] == (row[meta] if pd.notna(row[meta]) else None)
                            for meta in (expected_data_cols['meta'] + expected_data_cols['other'])) and
                        all(round(already_exists[num],2)==round(row[num],2)
                            for num in expected_data_cols['decimal'])
                ):
                    log.debug(f'detected data revision {row["Date"]:%Y-%m-%d} {row["email"]}\n'
                        + 'DB '     + (','.join(f'{col}:{already_exists[col]}' for col in expected_data_cols_all))
                        + ' email ' + (','.join(f'{col}:{row[col]}' for col in expected_data_cols_all))
                    )
                    has_data_revision = True

            if not already_exists or has_data_revision:
                data_table.insert(dict(
                    ParseTimestampUTC=parse_timestamp,
                    Attachment=attachment_name,
                    AttachmentHash=attachment_hash,
                    **{col: row[col] for col in expected_data_cols_all},
                    Revision=True if has_data_revision else None,
                ))

        db.commit()
    except:
        db.rollback()
        raise





def read_msgs():
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')
    inbox = outlook.GetDefaultFolder(constants.olFolderInbox)

    scan_recent_msgs = 30
    found_emails = []

    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem
    # To=mailitem.To,
    # From=mailitem.Sender.Name,
    # Subject=mailitem.Subject,
    # SentDt=f'{mailitem.SentOn:%Y-%m-%d %H:%M:%S}',
    # Received=f'{mailitem.ReceivedTime:%Y-%m-%d %H:%M:%S}',
    # Size=mailitem.Size,
    # Body=mailitem.Body
    # Attachments=','.join(a.FileName for a in mailitem.Attachments)
    def check_email(mailitem):
        log.debug(f'scanning email {mailitem.Sender.Name}|{mailitem.Subject:.40}|'
                  f'{mailitem.ReceivedTime:%Y-%m-%d %H:%M:%S}')
        if matched_attachments := match_mock_email(
                From=mailitem.Sender.Name,
                To=mailitem.To,
                Subject=mailitem.Subject,
                attached_filenames=[a.FileName for a in mailitem.Attachments]):
            log.debug(
                f'tagging email {mailitem.Sender.Name}|{mailitem.Subject:.40}|'
                f'{mailitem.ReceivedTime:%Y-%m-%d %H:%M:%S}')
            found_emails.append((mailitem, matched_attachments))

    # search within specific folders
    # - toplevel Inbox
    # - given foldername subfolder of Inbox
    # inbox.Items is a 1-index collection, ordered from oldest->newest
    log.debug(f'starting scan of Inbox of most recent {scan_recent_msgs}')
    for i in range(inbox.Items.Count, max(0, inbox.Items.Count - scan_recent_msgs), -1):
        check_email(inbox.Items[i])

    inbox_subfolder = 'python libs'
    if inbox_subfolder in (x.Name for x in inbox.Folders):
        log.debug(f'starting scan of Inbox subfolder [{inbox_subfolder}] of most recent {scan_recent_msgs}')
        for i in range(inbox.Folders[inbox_subfolder].Items.Count,
                       max(0, inbox.Folders[inbox_subfolder].Items.Count - scan_recent_msgs), -1):
            check_email(inbox.Folders[inbox_subfolder].Items[i])

    found_attachments = []
    for mailitem, matched_attachments_filenames in found_emails:
        for attachment_name in matched_attachments_filenames:
            log.debug(
                f'saving attachment {mailitem.Sender.Name}|{mailitem.Subject:.40}|'
                f'{mailitem.ReceivedTime:%Y-%m-%d %H:%M:%S}|{attachment_name}')
            temp_name = attachment_name.replace('.xls', f'_{mailitem.ReceivedTime:%Y-%m-%d_%H%M%S}.xls')
            temp_name = save_attachments_temp / temp_name
            if temp_name.exists():
                log.warning(f'overwriting existing temp file {temp_name}')
            mailitem.Attachments[attachment_name].SaveAsFile(temp_name)

            found_attachments.append((mailitem, attachment_name, temp_name))

    log.debug(f'found {len(found_attachments)} attachments to parse')
    for found_attachment in sorted(found_attachments, key=lambda tup: tup[0].ReceivedTime):
        log.debug(f'parsing data from {found_attachment[1]} from {found_attachment[0].Sender.Name}|'
                  f'{found_attachment[0].Subject:.40}|{found_attachment[0].ReceivedTime:%Y-%m-%d %H:%M}')
        parse_mock_attachment_to_db(*found_attachment)


if __name__ == '__main__':
    read_msgs()
