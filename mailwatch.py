#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 3 of the License, or
#  (at your option) any later version.
#
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software
#  Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

from errbot import BotPlugin, botcmd
import logging

import sys
PY2 = sys.version_info[0] == 2

import imaplib
from email.header import decode_header, make_header
import email, email.utils
import datetime

log = logging.getLogger('errbot.plugins.mailwatch')

class MailWatch(BotPlugin):
	""""Poll IMAP mailboxes and report new mails to specified chatrooms"""
	min_err_version = '1.6.0'
	_highest_uid = None # Highest UID we've encountered, so we know where we left off

	def activate(self):
		super(MailWatch, self).activate()
		
		if self.config is not None and set(("INTERVAL", "ACCOUNTS")) <= set(self.config):
			self._initial_poll = True
			self.start_poller(self.config['INTERVAL'], self.runpolls)
		else:
			log.info("Not starting MailWatch poller, plugin not configured")

	def get_configuration_template(self):
		return {'INTERVAL': 60, 'ACCOUNTS': [{'HOSTNAME': 'domain.tld', 'MAILBOX': 'INBOX', 'USERNAME': 'username', 'PASSWORD': 'password', 'ROOM': 'roomid@conference.domain.tld', 'SSL': True}]}

	def check_configuration(self, configuration):
		for i,item in enumerate(configuration['ACCOUNTS']):
			if 'MAILBOX' not in item:
				configuration['ACCOUNTS'][i]['MAILBOX'] = 'INBOX' 

	def runpolls(self):
		"""Polls all configured mailboxes"""
		assert 'ACCOUNTS' in self.config

		for account in self.config['ACCOUNTS']:
			self.poll(account['HOSTNAME'], account['MAILBOX'], account['USERNAME'], account['PASSWORD'], account['ROOM'])

	def poll(self, host, mailbox, user, passwd, room, ssl=True):
		"""Poll an IMAP mailbox"""
		log.info("Polling {0}@{1}".format(user, host))

		if 'seen' not in self.shelf.keys():
			seen = []
		else:
			seen = self.shelf['seen']
		if ssl:
			M = imaplib.IMAP4_SSL(host)
		else:
			M = imaplib.IMAP4(host)

		log.debug("IMAP LOGIN")
		code,message = M.login(user, passwd)
		log.debug("{0}: {1}".format(code, message))

		log.debug("IMAP SELECT: {0}".format(mailbox))
		M.select(mailbox)
		log.debug("{0}: {1}".format(code, message))
		log.debug("IMAP SEARCH")
		if self._highest_uid is None:
			search = '(SENTSINCE {})'.format((datetime.datetime.now() + datetime.timedelta(weeks=-1)).strftime('%d-%b-%Y'))
		else:
			search = '(UID {}:*)'.format(self._highest_uid)
		typ, data = M.search(None, search)
		log.debug("{0}: {1}".format(typ, data))

		for num in data[0].split():
			typ, data = M.fetch(num, '(RFC822.HEADER)')
			raw_mail = data[0][1]
			# raw_message is a bytestring which must be decoded to make it usable
			mail = email.message_from_string(raw_mail.decode("utf-8", "ignore"))
			if mail.get('Message-ID') not in seen:
				log.debug("New message: {0}".format(mail.get('Message-ID')))
				seen.append(mail.get('Message-ID'))

				message = 'New email arrived'
				for hdrname in ['From','To','Cc','Subject']:
					value = mail.get(hdrname) or None
					if value:
						if PY2:
							pairs = decode_header(value)
							hdrvalue = ' '.join([ unicode(t[0], t[1] or 'ASCII') for t in pairs ])
							message += "\n\t{}: {}".format(hdrname, hdrvalue)
						else:
							hi = make_header(decode_header(value))
							message += "\n\t{}: {}".format(hdrname, str(hi))

				self.send(room, message, message_type='groupchat')
			else:
				log.debug("Seen message: {0}".format(mail.get('Message-ID')))
			self._highest_uid = num.decode('ascii')
		M.close()
		M.logout()
		self.shelf['seen'] = seen
		self.shelf.sync()

