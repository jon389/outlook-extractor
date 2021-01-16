from pathlib import Path
from datetime import datetime
from dateutil import tz
import os, shutil, win32com.client, hashlib
from win32com.client import constants
import pandas as pd, numpy as np, xlwings as xw
# import dataset

from log_conf import logger as log, logs_folder

save_attachments_temp = logs_folder.parent / 'attachments_temp'
if not save_attachments_temp.exists():
    save_attachments_temp.mkdir()
for f in save_attachments_temp.iterdir():  # delete all files
    f.unlink()

try:
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
except AttributeError:
    shutil.rmtree(Path(os.environ['LOCALAPPDATA']) / 'gen_py', ignore_errors=True)
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')


expected_data_cols = dict(
    datetime='Date'.split(),
    meta='first_name last_name emailID gender ip_address'.split(),
    decimal='GST 1,PST-3,HST,JST,KST'.split(','),
    other='Comment'.split(),
)
expected_data_cols_all = lambda: [col for col_grp in expected_data_cols.values() for col in col_grp]


def get_parsed_attachments_table() -> pd.DataFrame:
    fname = save_attachments_temp.parent / f'parsed_attachments.csv'
    if fname.exists():
        return pd.read_csv(fname)
    return pd.DataFrame()


parsed_attachments = get_parsed_attachments_table()
parsed_attach_data = pd.DataFrame()


def match_mock_email(From: str, To: str, Subject: str, attached_filenames: list[str]) -> list[str]:
    if ('jon' in From.lower() and
            'jon' in To.lower() and
            'outlook' in Subject.lower() and 'extract' in Subject.lower()):
        matched_attachments = [a for a in attached_filenames
                               if ('mock' in a.lower() and 'attach' in a.lower()
                               and '.xls' in a.lower())]  # matches xls, xlsx, xlsm, xlsb
        return matched_attachments


def parse_mock_attachment_xls(attachment_name: str, attachment_temp_filename: Path) -> pd.DataFrame:

    data_sheets = 'Sheet1,Summary,'.split(',')

    with pd.ExcelFile(attachment_temp_filename) as xls:
        attach_data = pd.concat([
            pd.read_excel(xls,
                          sheet_name=sheet_name,
                          parse_dates=['Date'],
                          dtype={meta: str for meta in (expected_data_cols['meta'] + expected_data_cols['other'])},
                          )
            .assign(Comment=lambda d: (d.Comment.fillna('').astype(str) if 'Comment' in d else '')
                    )  # otherwise dataset creates as a float column
            for sheet_name in data_sheets if sheet_name in xls.sheet_names
        ])

    # check data consistency

    # rename some known column synonyms
    attach_data.rename(columns={
        'GST': 'GST 1',
        'PST': 'PST-3',
        'hst': 'HST',
        'JST 1': 'JST',
    }, inplace=True)
    if 'KST' not in attach_data:
        expected_data_cols['decimal'] = 'GST 1,PST-3,HST,JST'.split(',')

    #   eg. expected columns exist
    assert all(col in attach_data for col in expected_data_cols_all())

    # drop rows where emailID is na
    attach_data.dropna(subset='emailID gender ip_address'.split(), how='all', inplace=True)

    #   eg. Date col is parsed as a date
    assert 'datetime' in str(attach_data['Date'].dtype)
    # strip all strings, slow because applies func to each element
    attach_data.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # uppercase these fields
    attach_data['emailID gender ip_address'.split()] = \
        attach_data['emailID gender ip_address'.split()].apply(lambda x: x.str.upper())

    #   check that input data has unique (Date, emailID)
    if attach_data.duplicated('Date emailID'.split()).any():
        err_msg = f'data error, duplicate rows in {attachment_name}\n' \
                  + attach_data[attach_data.duplicated('Date emailID'.split(), keep=False)].to_string()

        log.warning(err_msg)
        #raise ValueError(err_msg)

        # aggregate duplicated rows
        attach_data = (attach_data
                       .groupby('Date emailID'.split(), as_index=False)
                       .agg(dict(**{c: lambda x: ','.join((str(x) if pd.notna(x) else '') for x in x.unique())
                                    for c in 'first_name last_name gender ip_address Comment'.split()},
                                 **{c: 'sum' for c in expected_data_cols['decimal']},)
                            )
                       )[expected_data_cols_all()]

    attach_data = (attach_data
                   # select all columns up to Comment rather than all columns [expected_data_cols_all()]
                   [attach_data.columns.tolist()[:attach_data.columns.tolist().index('Comment')+1]]
                   # change column names to non-reserved names
                   .rename(columns=lambda x: x.replace('-', '_').replace('.', '_').replace(' ', '_'))
                   .rename(columns={'Date': 'txDate'})
                   )

    return attach_data


def parse_each_mock_attachment(mailitem, attachment_name: str, attachment_temp_filename: Path):
    # ParsedEmails Table
    # ParseTimestampUTC
    # mailitem: From Subject ReceivedTime Body
    # Attachments -> list of attachment file names
    # SelectedAttachment -> file name of selected parsed xlsx
    # SelectedAttachmentHash

    # check if already parsed SelectedAttachment/SelectedAttachmentCheckSum
    global parsed_attachments, parsed_attach_data
    new_attach_data = []
    parse_timestamp = datetime.utcnow()  # datetime.now(tz.tzlocal())
    attachment_hash = hashlib.sha1(open(attachment_temp_filename, 'rb').read()).hexdigest()

    already_exists = parsed_attachments.query(
        f'Attachment=="{attachment_name}" & AttachmentHash=="{attachment_hash}"'
    ) if not parsed_attachments.empty else pd.DataFrame()

    if not already_exists.empty:
        already_exists = already_exists.iloc[0]
        log.info(f'ignoring {attachment_name} from {mailitem.Sender.Name}|{mailitem.Subject:.40}|'
                    f'{mailitem.ReceivedTime:%Y-%m-%d %H:%M}'
                 f' already in ParsedEmails table '
                 f'from {already_exists["From"]}|{already_exists["Subject"]:.40}|'
                 f'{already_exists["Received"]:%Y-%m-%d %H:%M}')
        return

    try:
        attach_data = parse_mock_attachment_xls(attachment_name, attachment_temp_filename)
        if attach_data.empty:
            return

        for idx, row in attach_data.iterrows():
            has_data_revision = False
            if not parsed_attach_data.empty:
                data_query = (parsed_attach_data['txDate' ] == row['txDate' ]) & \
                             (parsed_attach_data['emailID'] == row['emailID'])
                if data_query.any():
                    data_exist = parsed_attach_data[data_query]\
                        .sort_values('ParseTimestampUTC', ascending=False).iloc[0]

                    # log.debug(f'data row txDate {row["txDate"]:%Y-%m-%d} {row["emailID"]} already exist')
                    if not ( all(data_exist[meta] == row[meta]
                                 for meta in [x for x in attach_data.select_dtypes(exclude='number').columns
                                              if x not in 'ip_address Comment'.split()])
                         and all(round(data_exist[num], 2) == round(row[num], 2)
                                 for num  in attach_data.select_dtypes(include='number').columns)
                    ):
                        log.info(f'detected data revision {row["txDate"]:%Y-%m-%d} {row["emailID"]}'
                            + '\nDB   :' + (','.join(f'{col}:{data_exist[col]}' for col in attach_data.columns))
                            + '\nemail:' + (','.join(f'{col}:{row[col]}'        for col in attach_data.columns)) )
                        has_data_revision = True

            if parsed_attach_data.empty or not data_query.any() or has_data_revision:
                new_attach_data.append(dict(
                    ParseTimestampUTC=parse_timestamp,
                    Attachment=attachment_name,
                    AttachmentHash=attachment_hash,
                    **{col: row[col] for col in attach_data.columns},
                    Revision=True if has_data_revision else None,
                ))

        parsed_attach_data = pd.concat([parsed_attach_data, pd.DataFrame(new_attach_data)], ignore_index=True)
        parsed_attachments = parsed_attachments.append(dict(
            ParseTimestampUTC=parse_timestamp,
            To=mailitem.To,
            From=mailitem.Sender.Name,
            Subject=mailitem.Subject,
            Received=mailitem.ReceivedTime.replace(tzinfo=None),
            Size=mailitem.Size,
            Body=mailitem.Body,
            EmailAttachments=','.join(a.FileName for a in mailitem.Attachments),
            Attachment=attachment_name,
            AttachmentHash=attachment_hash,
        ), ignore_index=True)

    except:
        raise


def read_msgs():
    inbox = outlook.GetNamespace('MAPI').GetDefaultFolder(constants.olFolderInbox)

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
        try:
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
        except:
            pass

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
        parse_each_mock_attachment(*found_attachment)  # (mailitem, attachment_name, temp_name)


if __name__ == '__main__':
    read_msgs()

    parsed_attachments.to_csv(
        save_attachments_temp.parent / f'parsed_attachments_{datetime.now():%Y-%m-%d_%H%M%S}.csv', index=False)

    # only return most recent revision (by ParseTimestampUTC), note parsing order is done by ReceivedTime order
    parsed_attach_data.loc[parsed_attach_data.groupby('txDate emailID'.split())['ParseTimestampUTC'].idxmax()]\
        .to_csv(save_attachments_temp.parent / f'parsed_data_{datetime.now():%Y-%m-%d_%H%M%S}.csv',
                index=False)

    summary_file = Path(r'C:\temp') / f'Summary_data_{datetime.now():%Y-%m-%d_%H%M%S}.csv'
    # save to file only most recent revision of each row
    parsed_attach_data.loc[parsed_attach_data.groupby('txDate emailID'.split())['ParseTimestampUTC'].idxmax()]\
        .to_csv(summary_file, index=False)
    log.info(summary_file)