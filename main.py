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
        mailitem = inbox.Items[i]
        # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem
        deets = dict(To=mailitem.To,
                     From=mailitem.Sender.Name,
                     Subject=mailitem.Subject,
                     # SentDt=f'{mailitem.SentOn:%Y-%m-%d %H:%M:%S}',
                     Received=f'{mailitem.ReceivedTime:%Y-%m-%d %H:%M:%S}',
                     Size=mailitem.Size
                     )
        log.info(deets)
        recent_5_msgs.append(deets)


if __name__ == '__main__':
    read_msgs()