#
#  App.rb
#  Ent Menu
#
#  Created by Steve Loveless on 3/8/09.
#  Copyright (c) 2009 Pelco. All rights reserved.
#

require 'osx/cocoa'
include OSX
OSX.require_framework 'ScriptingBridge'

class App < NSObject
	@status_item
	@menu
	@entourage
	@ent_accts
	@update_interval

	#----------------------------------------------------------------------------
	# Method:			initialize
	#
	# Purpose:		
	#----------------------------------------------------------------------------
	def initialize()
		@entourage = SBApplication.applicationWithBundleIdentifier_("com.microsoft.Entourage")
		@update_interval = 300.0
		entAccounts()
	end

	#----------------------------------------------------------------------------
	# Method:			applicationDidFinishLaunching
	#
	# Purpose:		
	#----------------------------------------------------------------------------
  def applicationDidFinishLaunching(aNotification)
		# Init status bar
    statusbar = NSStatusBar.systemStatusBar
    @status_item = statusbar.statusItemWithLength(NSVariableStatusItemLength)
		
    #image = NSImage.alloc.initWithContentsOfFile("Entourage_mac_2008_icon.png")
    #status_item.setImage(image)
		
		# Create menu with the mail count
		postFirstCount()
		
		# Update the count at the set interval
		@timer = NSTimer.scheduledTimerWithTimeInterval_target_selector_userInfo_repeats_(@update_interval, self, :updateCount, nil, true)
  end

	#----------------------------------------------------------------------------
	# Method:			applicationWillTerminate
	#
	# Purpose:		
	#----------------------------------------------------------------------------  
  def applicationWillTerminate(aNotification)
    puts "Goodbye, cruel world!"
  end

	#----------------------------------------------------------------------------
	# Method:			mailCount
	#
	# Purpose:		Gets the number of emails from the first Exch. inbox
	#----------------------------------------------------------------------------
  def mailCount
		#ent_accts.each do |acct|
			#puts acct.name
			#puts acct.emailAddress
			# set inboxFolder to (inbox folder of Exchange account acctName)
		#	inbox = acct.inboxFolder.get
			# set totalCount to (unread message count of theAcct)
		#	emailcount = inbox.unreadMessageCount
			#puts count
		#end
		inbox = @ent_accts[0].inboxFolder.get
		@emailcount = inbox.unreadMessageCount
		puts @emailcount
		return @emailcount.to_s
  end
	
	#----------------------------------------------------------------------------
	# Method:			postFirstCount
	#
	# Purpose:		
	#----------------------------------------------------------------------------
	def postFirstCount
	  text = NSString.alloc.initWithString(mailCount())
    @status_item.setTitle(text)
    initMenu(@status_item)
	end

	#----------------------------------------------------------------------------
	# Method:			updateCount
	#
	# Purpose:		
	#----------------------------------------------------------------------------
	def updateCount(sender)
	  text = NSString.alloc.initWithString(mailCount())
    @status_item.setTitle(text)
    updateMenu(@status_item)
	end

	#----------------------------------------------------------------------------
	# Method:			initMenu
	#
	# Purpose:		
	#----------------------------------------------------------------------------
	def initMenu(container)
    @menu = NSMenu.alloc.init
    container.setMenu(@menu)
		menu_item = @menu.addItemWithTitle_action_keyEquivalent("#{mailCount} message(s)" , nil, "")
    menu_item = @menu.addItemWithTitle_action_keyEquivalent("Check for new", "updateCount:", '')
		menu_item = @menu.addItemWithTitle_action_keyEquivalent("Bring Entourage to front", "activateEnt:", '')
		menu_item = @menu.addItemWithTitle_action_keyEquivalent("Compose new message", "newMessage:", '')
    menu_item = @menu.addItemWithTitle_action_keyEquivalent("Quit", "terminate:", '')
    menu_item.setKeyEquivalentModifierMask(NSCommandKeyMask)
    menu_item.setTarget(NSApp)
	end

	#----------------------------------------------------------------------------
	# Method:			updateMenu
	#
	# Purpose:		
	#----------------------------------------------------------------------------
	def updateMenu(container)
    container.setMenu(@menu)
		@menu.removeItemAtIndex_(0)
		menu_item = @menu.insertItemWithTitle_action_keyEquivalent_atIndex("#{mailCount} message(s)" , nil, "",0)
	end

	#----------------------------------------------------------------------------
	# Method:			activateEnt
	#
	# Purpose:		
	#----------------------------------------------------------------------------
	def activateEnt(sender)
		@entourage.activate
	end

	#----------------------------------------------------------------------------
	# Method:			entAccounts
	#
	# Purpose:		Gets list to Exch. accounts
	#----------------------------------------------------------------------------
	def entAccounts
		@ent_accts = @entourage.ExchangeAccounts
	end
	
	#----------------------------------------------------------------------------
	# Method:			newMessage
	#
	# Purpose:		
	#----------------------------------------------------------------------------	
	def newMessage(sender)
		thescript = "tell application \"Microsoft Entourage\"
	set theMessage to make new outgoing message
	open theMessage
end tell"
		NSAppleScript.alloc.initWithSource_(thescript).executeAndReturnError_(nil)
		@entourage.activate
	end
end


NSApplication.sharedApplication
NSApp.setDelegate(App.alloc.init)
NSApp.run
