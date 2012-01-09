# iMapBox.py
#
# Module to easily sift through imap emailboxes
# Copyright (c) 2011, Jerry Neumann
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated 
# documentation files (the "Software"), to deal in the Software without restriction, including without limitation 
# the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, 
# and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions 
# of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
# TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
# THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF 
# CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
# DEALINGS IN THE SOFTWARE.
#
# To add:
#    attachments filter, flag and part
#    flags filters (read, unread, etc.)

import imaplib, datetime, email, string, re

globmap = {'to':'To',
			'from':'From',
			'cc':'Cc',
			'bcc':'Bcc',
			'date':'Date',
			'time':'Date',
			'subject':'Subject',
			'text':''}
#			'attachment':''}

globmonths=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def multiton(cls):
	# set up multiton decorator
	instances = {}
	def getinstance(id1,id2):
		if (id1,id2) not in instances:
			instances[(id1,id2)] = cls(id1,id2)
		return instances[(id1,id2)]	
	return getinstance
	
def getParts(unparsed,obj):
	parseparts = email.message_from_string(unparsed)
	for mypart in globmap:
		obj.parts[mypart] = parseparts.get(globmap[mypart],None)
	
	# parse out date and time
	year,month,day,hour,minute,seconds,x,x,dst = email.utils.parsedate(obj.parts['date'])
	obj.parts['date'] = datetime.date(year,month,day)
	obj.parts['time'] = datetime.time(hour, minute, seconds)
		
	# set attachment?--how to know if there is an attachment?
	# self.parts['attachment?'] = False
	# if ????:
	#	self.parts['attachment?'] = True

class IMapBox(object):
	"""Opens an IMAP account.
	Use:  acct = IMapBox(server, username, password, [port], [priority (headers/both)])
	Methods:
	   - list(): returns a list of mailboxes
	   - __getitem__(mailbox): returns a MsgBox representing the mailbox
	   - __del__(): logs out, closes connection (not being used, breaks chaining.)"""
	
	def __init__(self, server, username, password, port=993,priority='headers'):
		if priority not in ['headers','text','all']: raise ValueError("priority must be 'headers', 'text' or 'all'.")
		list_resp_re = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')
		self.priority=priority
		self.svr = imaplib.IMAP4_SSL(server,port)
		self.svr.login(username,password)
		self.status,resp=self.svr.list()
		self.mailboxes = [list_resp_re.match(mb).groups()[2].strip('\"\'') for mb in resp]
	
	def __getitem__(self, mailbox):
		if mailbox not in self.mailboxes:
			raise KeyError(mailbox)
		
		status, strlen = self.svr.select(mailbox)	
		return MsgBox(self,mailbox)
		
	def list(self):
		return self.mailboxes

#	not using __del__ to close connection because chaining causes
#	the IMapBox object to be garbage collected. Need to fix.

#	def __del__(self):
#		self.svr.close()
#		self.svr.logout()
		


class MsgBox(object):
	"""
	A dictionary of messages. Keys=msg ids, values=BoxMsgs.
	Use: MsgBox(MapBox,srch="search string")
	        Creates a MsgBox object on mailbox, qualified with search string

	Methods:
	Dict methods:
	   - __getitem__ : obj[msgid] returns Msg object for the msgid
	   - __len__: returns number of msgs
	   - __keys__: returns list of msgids
	   - __items__: returns list of Msg objects
	   - __contains__(msgid): return True/False for msgid in box
	   - get(x,y): __getitem__(x) or y if no x
	   - __iter__(): iterate through msgids
	   
	Filter methods:
	   Returns a MsgBox filtered through IMAP search
	   - .to(x): filters for messages to x
	   - .frm(x)
	   - .cc(x)
	   - .bcc(x)
	   - .dates(x,[y]): filters for dates between x and y, if no y then just msgs on date x
	      x and y are datetime.dates
	   - .today(): messages with today's date
	   - .subject(x)
	   - .times(x,y): filters for time between x and y. NOT YET IMPLEMENTED.
	
	"""
	
	def __init__(self,calling_box,mailbox,**kwargs):
		self.mb_name = mailbox
		self.msglist = {}
		self.svr = calling_box.svr
		self.priority = calling_box.priority
		
		# check if srch arg is present, if it is, augment srch and return a new MsgBox
		self.srch = kwargs.get('srch','')
		
	def _getMsgs(self):
		# takes the srch string, fetches the msgids and creates the dictionary of {msgid:BoxMsg(msgid),. . .}
		if not self.msglist:
			if not self.srch:
				msgs = self.svr.search(None,'(ALL)')
			else:
				msgs = self.svr.search(None,'('+self.srch.strip()+')')
			if msgs[0] == "OK":
				msgids = msgs[1][0].split()
				for v in msgids:
					self.msglist[v] = BoxMsg(self,v)
	
	def _fetchMsgs(self):
		# fetch all the msg headers in the current msglist and stuff their info into
		# the right BoxMsg objects. Only used for __iter__, values() and items() right now.
		
		# fetch all msgs in self.msglist at once
		# for each message, parse it and load parts into BoxMsgs
		
		#! Need to fix because priority effects knowing whether or not a message is already fetched
		
		notftchd = [i for i in self.msglist.keys() if not self.msglist[i].hdr_fetched]
		ftchlist = ",".join(notftchd)
		if ftchlist:
			if self.priority == 'all':
				ftchtype = '(BODY.PEEK[])'
			elif self.priority == 'headers':
				ftchtype = '(BODY.PEEK[HEADER])'
			else:
				ftchtype = '(BODY.PEEK[TEXT])'
			ftchd = self.svr.fetch(ftchlist,ftchtype)
			ftchdmsgs = ftchd[1][0::2]
			for msg in ftchdmsgs:
				m1,m2=msg
				msgid=m1.split()[0]
				getParts(m2,self.msglist[msgid])
				self.msglist[msgid].hdr_fetched = True
		
	def __len__(self):
		self._getMsgs()
		return len(self.msglist)
		
	def keys(self):
		self._getMsgs()
		return self.msglist.keys()
		
	def items(self):
		self._getMsgs()
		self._fetchMsgs()
		return self.msglist.items()
		
	def values(self):
		self._getMsgs()
		self._fetchMsgs()
		return self.msglist.values()
		
	def __getitem__(self,msgid):
		self._getMsgs()
		return self.msglist[msgid]
		
	def __contains__(self,msgid):
		return msgid in self.msglist
		
	def __iter__(self):
		self._getMsgs()
		self._fetchMsgs()
		return iter(self.msglist.keys())
		
	def get(self,a,b):
		self._getMsgs()
		if a in self.msglist: return self.msglist[a]
		return b
		
	def __add__(self,other):
		return MsgBox(self,self.mb_name,srch=''.join(["OR (",self.srch.strip(),") (",other.srch.strip(),") "]))
	
	def __sub__(self,other):
		return MsgBox(self,self.mb_name,srch=''.join(["(",self.srch.strip(),") NOT (",other.srch.strip(),") "]))

	def __neg__(self):
		return MsgBox(self,self.mb_name,srch=''.join(["NOT (",self.srch.strip(),") "]))
	
	def _makedate(self,d):
		# date args are datetimes, but need to be in form 30-May-1966 or 05-Jan-2010
		return "%s-%s-%s" % (d.day,globmonths[d.month-1],d.year)
	
	def to(self,augsrch): return MsgBox(self,self.mb_name,srch=''.join([self.srch,"TO ",augsrch," "]))
	def frm(self,augsrch): return MsgBox(self,self.mb_name,srch=''.join([self.srch,"FROM ",augsrch," "]))
	def cc(self,augsrch): return MsgBox(self,self.mb_name,srch=''.join([self.srch,"CC ",augsrch," "]))
	def bcc(self,augsrch): return MsgBox(self,self.mb_name,srch=''.join([self.srch,"BCC ",augsrch," "]))
	def subject(self,augsrch): return MsgBox(self,self.mb_name,srch=''.join([self.srch,"SUBJECT ",augsrch," "]))
	def today(self): return MsgBox(self,self.mb_name,srch=''.join([self.srch,"ON ",self._makedate(datetime.datetime.now())," "]))
	def dates(self,arg1,arg2=None):
		# two args: a begin and an end; one arg: messages only on that date
		if arg2:
			return MsgBox(self,self.mb_name,srch=''.join([self.srch,"SINCE ",self._makedate(arg1)," BEFORE ",self._makedate(arg2)," "]))
		else:
			return MsgBox(self,self.mb_name,srch=''.join([self.srch,"ON ",self._makedate(arg1)," "]))
	
@multiton
class BoxMsg(object):
	"""
	Each object is a lazily evaluated message. Initialized by a call to BoxMsg(MsgBox, msgid).
	BoxMsg is a multiton, so there is only one of each mailbox/msgid.
	
	Use:
	msg = BoxMsg(MsgBox,msgid)
	A call to msg[part] returns that part of the message. part can be:
	   - to: returns list of recipients
	   - from: returns sender
	   - cc: returns list of cc'd
	   - bcc: returns list of bcc'd
	   - date: returns date sent
	   - time: returns time sent
	   - subject: returns subject
	   - text: returns text"""	   
#	   - attachment?: returns True if attachment, otherwise False -- Not Yet Implemented
	
	def __init__(self,msgbox,msgid):
		self.svr = msgbox.svr
		self.priority=msgbox.priority
		self.id = msgid
		self.parts = {}
		self.hdr_fetched = False
		self.txt_fetched = False
		
	def __getitem__(self,part):
		if part not in globmap.keys():
			raise KeyError(part)			
		if not self.hdr_fetched and (part!='text' or self.priority!='text'):
			hdr = self.svr.fetch(self.id,'(BODY.PEEK[HEADER])')
			getParts(hdr[1][0][1],self)
			self.hdr_fetched=True		
		if not self.txt_fetched and (part=='text' or self.priority!='headers'):
			txt = self.svr.fetch(self.id,'(BODY.PEEK[TEXT])')
			self.parts['text']=txt[1][0][1]
			self.txt_fetched = True			
		return self.parts[part]