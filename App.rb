#
#  App.rb
#  Ent Menu
#
#  Created by Steve Loveless on 3/8/09.
#  Copyright (c) 2009 Steve Loveless. All rights reserved.
#

require 'osx/cocoa'
include OSX
OSX.require_framework 'ScriptingBridge'

class App < NSObject
  @status_item
  @menu
  @entourage
  @exch_accounts
  @imap_accounts
  @update_interval

  #----------------------------------------------------------------------------
  # Method:			initialize
  #
  # Purpose:		
  #----------------------------------------------------------------------------
  def initialize()
    @entourage = SBApplication.applicationWithBundleIdentifier_("com.microsoft.Entourage")
    @update_interval = 300.0
    exchAccounts()
    imapAccounts()
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
  # Purpose:		Gets the number of emails from the specified account type
  #----------------------------------------------------------------------------
  def mailCount count_type=nil
    # Init the email count to start fresh every time this is called
    email_count = 0

    # Get the total count of emails in the Exchange inboxes
    exch_count = exchangeMailCount
    email_count += exch_count
    
    # Get the total count of emails in the IMAP inboxes
    imap_count = imapMailCount
    email_count += imap_count

    puts email_count
    return email_count.to_s
  end


  #----------------------------------------------------------------------------
  # Method:			exchangeMailCount
  #
  # Purpose:		Gets the number of emails from each Exch. inbox
  #----------------------------------------------------------------------------
  def exchangeMailCount
    # Init the email count to start fresh every time this is called
    email_count = 0

    # Get the total count of emails in the Exchange inboxes
    exchange_inbox = Array.new
    count = 0
    @exch_accounts.each do |acct|
      exchange_inbox[count] = @exch_accounts[count].inboxFolder.get
      email_count += exchange_inbox[count].unreadMessageCount
    end

    return email_count
  end
	
  #----------------------------------------------------------------------------
  # Method:			imapMailCount
  #
  # Purpose:		Gets the number of emails from the specified account type
  #----------------------------------------------------------------------------
  def imapMailCount
    # Init the email count to start fresh every time this is called
    email_count = 0

    # Get the total count of emails in the IMAP inboxes
    imap_inbox = Array.new
    count = 0
    @imap_accounts.each do |acct|
      imap_inbox[count] = @imap_accounts[count].IMAPInboxFolder.get
      email_count += imap_inbox[count].unreadMessageCount
      count += 1
    end

    return email_count
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
    if @entourage.isRunning
      text = NSString.alloc.initWithString(mailCount())
      @status_item.setTitle(text)
      updateMenu(@status_item)
    else
      puts "Entourage not running anymore.  Quitting."
      NSApplication.sharedApplication.terminate(self)
    end
  end

  #----------------------------------------------------------------------------
  # Method:			initMenu
  #
  # Purpose:		
  #----------------------------------------------------------------------------
  def initMenu(container)
    @menu = NSMenu.alloc.init
    container.setMenu(@menu)
    menu_item = @menu.addItemWithTitle_action_keyEquivalent("Total: #{mailCount} message(s)" , nil, "")
    menu_item = @menu.addItemWithTitle_action_keyEquivalent("Exchange: #{exchangeMailCount} message(s)" , nil, "")
    menu_item = @menu.addItemWithTitle_action_keyEquivalent("IMAP: #{imapMailCount} message(s)" , nil, "")
    menu_item = @menu.addItemWithTitle_action_keyEquivalent("--------------------" , nil, "")
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
    menu_item = @menu.insertItemWithTitle_action_keyEquivalent_atIndex("Total: #{mailCount} message(s)" , nil, "",0)
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
  # Method:		    quitEnt	
  #
  # Purpose:		
  #----------------------------------------------------------------------------
  def quitEnt(sender)
    @entourage.quit
  end

  #----------------------------------------------------------------------------
  # Method:			exchAccounts
  #
  # Purpose:		Gets list to Exch. accounts
  #----------------------------------------------------------------------------
  def exchAccounts
    @exch_accounts = @entourage.ExchangeAccounts
    puts @exch_accounts
  end
	
  #----------------------------------------------------------------------------
  # Method:			imapAccounts
  #
  # Purpose:		Gets list to IMAP accounts
  #----------------------------------------------------------------------------
  def imapAccounts
    @imap_accounts = @entourage.IMAPAccounts
    puts @imap_accounts
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
