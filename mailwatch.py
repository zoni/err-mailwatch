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

import imaplib
import email, email.utils, email.header
import datetime

class MailWatch(BotPlugin):
	""""Poll IMAP mailboxes and report new mails to specified chatrooms"""
	min_err_version = '1.6.0'
	_initial_poll = True # True if poll hasn't run yet since plugin activation
	_highest_uid = None # Highest UID we've encountered, so we know where we left off

	def activate(self):
		super(MailWatch, self).activate()
		
		if self.config is not None and set(("INTERVAL", "ACCOUNTS")) <= set(self.config):
			self._initial_poll = True
			self.start_poller(self.config['INTERVAL'], self.runpolls)
		else:
			logging.info("Not starting MailWatch poller, plugin not configured")

	def get_configuration_template(self):
		return {'INTERVAL': 60, 'ACCOUNTS': [{'HOSTNAME': 'domain.tld', 'USERNAME': 'username', 'PASSWORD': 'password', 'ROOM': 'roomid@conference.domain.tld', 'SSL': True}]}

	def runpolls(self):
		"""Polls all configured mailboxes"""
		assert 'ACCOUNTS' in self.config

		for account in self.config['ACCOUNTS']:
			self.poll(account['HOSTNAME'], account['USERNAME'], account['PASSWORD'], account['ROOM'])

	def poll(self, host, user, passwd, room, ssl=True):
		"""Poll an IMAP mailbox"""
		logging.info("Polling {0}@{1}".format(user, host))
		if 'seen' not in self.shelf.keys():
			seen = []
		else:
			seen = self.shelf['seen']
		if ssl:
			M = imaplib.IMAP4_SSL(host)
		else:
			M = imaplib.IMAP4(host)

		logging.debug("IMAP LOGIN")
		code,message = M.login(user, passwd)
		logging.debug("{0}: {1}".format(code, message))

		logging.debug("IMAP SELECT")
		M.select()
		logging.debug("{0}: {1}".format(code, message))
		logging.debug("IMAP SEARCH")
		if self._initial_poll:
			self._initial_poll = False
			# UID's *might* have changed since last time we checked, so start all over, looking only at mail sent in the last week
			search = '(SENTSINCE {})'.format((datetime.datetime.now() + datetime.timedelta(weeks=-1)).strftime('%d-%b-%Y'))
		else:
			search = 'UID {}:*'.format(self._highest_uid)
		typ, data = M.search(None, search)
		logging.debug("{0}: {1}".format(typ, data))

		for num in data[0].split():
			typ, data = M.fetch(num, '(RFC822.HEADER)')
			raw_mail = data[0][1]
			# raw_message is a bytestring which must be decoded to make it usable
			mail = email.message_from_string(raw_mail.decode("utf-8", "ignore"))
			if mail.get('Message-ID') not in seen:
				logging.debug("New message: {0}".format(mail.get('Message-ID')))
				seen.append(mail.get('Message-ID'))
				message = "New email arrived"
				message += "\n\tFrom: {0}".format(email.header.decode_header(mail.get('from'))[0][0])
				message += "\n\tTo: {0}".format(email.header.decode_header(mail.get('to'))[0][0])
				message += "\n\tCc: {0}".format(email.header.decode_header(mail.get('cc'))[0][0])
				message += "\n\tSubject: {0}".format(email.header.decode_header(mail.get('subject'))[0][0])
				self.send(room, message, message_type='groupchat')
			else:
				logging.debug("Seen message: {0}".format(mail.get('Message-ID')))
			self._highest_uid = num
		M.close()
		M.logout()
		self.shelf['seen'] = seen
		self.shelf.sync()

