#!usr/bin/env python
# coding: latin-1

"""
Bloomberg.

Bloomberg API wrapper to make syncronous requests.

Functions:
    bdp: reference request, similar to Excel BDP.
    bdh: historical request, similar to Excel BDH.
"""

import datetime as dt

import comtypes.client as cc # http://sourceforge.net/projects/comtypes/
import pandas as pan # http://code.google.com/p/pandas/

#-------------------------------------------------------------------------------
# Constants
#-------------------------------------------------------------------------------

UNKNOWN = -1
ADMIN = 1
SESSION_STATUS = 2
SUBSCRIPTION_STATUS = 3
REQUEST_STATUS = 4
RESPONSE = 5
PARTIAL_RESPONSE = 6
SUBSCRIPTION_DATA = 8
BLPSERVICE_STATUS = 9
TIMEOUT = 10
AUTHORIZATION_STATUS = 11
RESOLUTION_STATUS = 12
PUBLISHING_DATA = 13
TOPIC_STATUS = 14
TOKEN_STATUS = 15

#-------------------------------------------------------------------------------
# Errors handling
#-------------------------------------------------------------------------------

class BloombergError(Exception):
    """ No response from Bloomberg. """
    
    pass

#-------------------------------------------------------------------------------
# Bloomberg request functions
#-------------------------------------------------------------------------------

def bdp(sec_list, fld_list, verbose=False, **kwargs):
    """ Sends a reference request to Bloomberg.
    Parameters:
        sec_list: tuple or list of valid Bloomberg tickers.
        fld_list: tuple or list of valid Bloomberg fields.
        verbose: boolean to log Bloomberg response messages (default: False)
        **kwargs: any valid parameter.
    Returns a pandas.DataFrame object.
    """
    session = cc.CreateObject('blpapicom.Session')
    try:
        session.Start()
        session.OpenService('//blp/refdata')
        refdataservice = session.GetService('//blp/refdata')
        req = refdataservice.CreateRequest('ReferenceDataRequest')
        for s in sec_list:
            req.GetElement('securities').AppendValue(s)
        for f in fld_list:
            req.GetElement('fields').AppendValue(f)
        session.SendRequest(req)
        loop = True
        response = {}
        while loop:
            event = session.NextEvent()
            if verbose:
                print('Bloomberg event: %s' % event.EventType)
            iterator = event.CreateMessageIterator()
            iterator.Next()
            message = iterator.Message
            if verbose:
                print('Bloomberg message: %s' % message.MessageTypeAsString)
            if event.EventType == RESPONSE:
                num_securities = message.GetElement('securityData').NumValues
                for i in range(num_securities):
                    security = message.GetElement('securityData').GetValue(i)
                    name = security.GetElement('security').Value
                    response[name] = {}
                    fields = security.GetElement('fieldData')
                    for n in range(fields.NumElements):
                        field = fields.GetElement(n)
                        response[name][field.Name] = field.Value
                loop = False
        tempdict = {}
        for r in response:
            tempdict[r] = pan.Series(response[r])
        data = pan.DataFrame(tempdict)
        if data:
            return(data)
        else:
            raise BloombergError('No response from Bloomberg. Please check \
request arguments: tickers and fields must be tuples or lists.')
    except BloombergError:
        raise
    finally:
        session.Stop()
        iterator = None
        message = None
        event = None
        session = None

def bdh(sec_list, fld_list, start_date,
    end_date=dt.date.today().strftime('%Y%m%d'), periodicity='DAILY',
    verbose=False, **kwargs):
    """ Sends a historical request to Bloomberg.
    Parameters:
        sec_list: tuple or list of valid Bloomberg tickers.
        fld_list: tuple or list of valid Bloomberg fields.
        start_date: string formatted YYYYMMDD.
        end_date: string formatted YYYYMMDD (default = Today()).
        periodicity: string (default: DAILY).
        verbose: boolean to log Bloomberg response messages (default: False)
        **kwargs: any valid parameter.
    Returns a panda.Panel object.
    """
    session = cc.CreateObject('blpapicom.Session')
    try:
        session.Start()
        session.OpenService('//blp/refdata')
        refdataservice = session.GetService('//blp/refdata')
        req = refdataservice.CreateRequest('HistoricalDataRequest')
        for s in sec_list:
            req.GetElement('securities').AppendValue(s)
        for f in fld_list:
            req.GetElement('fields').AppendValue(f)
        req.Set('periodicitySelection', periodicity)
        req.Set('startDate', start_date)
        req.Set('endDate', end_date)
        session.SendRequest(req)
        loop = True
        response = {}
        while loop:
            event = session.NextEvent()
            if verbose:
                print('Bloomberg event: %s' % event.EventType)
            iterator = event.CreateMessageIterator()
            iterator.Next()
            message = iterator.Message
            if verbose:
                print('Bloomberg message: %s' % message.MessageTypeAsString)
            if event.EventType == RESPONSE or event.EventType == \
                PARTIAL_RESPONSE:
                security_data = message.GetElement('securityData')
                name = security_data.GetElement('security').Value
                response[name] = {}
                field_data = security_data.GetElement('fieldData')
                for i in range(field_data.NumValues):
                    fields = field_data.GetValue(i)
                    for n in range(1, fields.NumElements):
                        date = fields.GetElement(0).Value
                        field = fields.GetElement(n)
                        try:
                            response[name][field.Name][date] = field.Value
                        except KeyError:
                            response[name][field.Name] = {}
                            response[name][field.Name][date] = field.Value
            if event.EventType == RESPONSE:
                loop = False
        tempdict = {}
        for r in response:
            td = {}
            for f in response[r]:
                td[f] = pan.Series(response[r][f])
            tempdict[r] = pan.DataFrame(td)
        data = pan.Panel(tempdict)
        if data[[d for d in data][0]]:
            return(data)
        else:
            raise BloombergError('No response from Bloomberg. Please check \
request arguments: tickers and fields must be tuples or lists.')
    except BloombergError:
        raise
    finally:
        session.Stop()
        iterator = None
        message = None
        event = None
        session = None

#-------------------------------------------------------------------------------
# Test function
#-------------------------------------------------------------------------------

def test():
    """ Test function. """
    secs = ('C US Equity', 'GOOG US Equity')
    st_flds = ('NAME', 'PX_LAST')
    ref = bdp(secs, st_flds)
    print('A Bloomberg reference request example')
    print(ref)
    print('')
    his_flds = ('BEST_TARGET_PRICE', 'PX_LAST')
    beg = (dt.date.today() - dt.timedelta(30)).strftime('%Y%m%d')
    hist = bdh(secs, his_flds, beg)
    print('A Bloomberg historical request example')
    for h in hist:
        print(h)
        print(hist[h])
        
if __name__ == '__main__':
    test()