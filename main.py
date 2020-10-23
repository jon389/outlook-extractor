import pandas as pd, numpy as np, xlwings as xw
import win32com.client
from win32com.client import constants
from datetime import datetime
from log_conf import logger as log


def read_msgs():
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')
    inbox = outlook.GetDefaultFolder(constants.olFolderInbox)

    recent_5_msgs = []
    # inbox.Items is a 1-index collection, ordered from oldest->newest
    for i in range(inbox.Items.Count, inbox.Items.Count-5, -1):
        message = inbox.Items[i]
        deets = dict(SentDt=f'{message.SentOn:%Y-%m-%d %H:%M:%S}',
                     Sender=message.Sender.Name,
                     To=message.To,
                     Subject=message.Subject
                     )
        log.info(deets)
        recent_5_msgs.append(deets)


if __name__ == '__main__':
    read_msgs()