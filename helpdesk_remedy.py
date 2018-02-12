import win32gui, time, SendKeys, re, win32clipboard, win32con, string, win32com.client
def WinEnumHandler(hwnd, resultList):
   resultList.append((hwnd, win32gui.GetWindowText(hwnd)))
 
topWindows = []  #--> array to store all (hwnd, name) values of all top-windows
win32gui.EnumWindows(WinEnumHandler, topWindows)
   #--> will put all (hwnd, name) values of top-windows into topWindows array.
 # Find desired window based on <string> within it's name.
count = 0
stuff = ""
for row in topWindows:
   x = re.search('Certificate Management Center', row[1], 2)
   if x:
      myWindow = [row[0],row[1],win32gui.GetClassName(row[0])]
      count = count + 1
      stuff = stuff + myWindow[1]
      #print myWindow[1]
   elif x == False:
      print "It appears you do not have TSAdmin loaded in a webbrowser."
      break
      #print count   
   
#print stuff
if count >= 2:
	c = input("Which browser do you want me to use? Type 1 for Firefox, 2 for IE: ")

	if c == 1:
		x = win32gui.FindWindow(None, "Certificate Management Center - Mozilla Firefox") #get hwnd given the window title
		win32gui.SetForegroundWindow(x) #bring the window with x hwnd to the front
		SendKeys.SendKeys('^{u}') #view page source in firefox
		#time.sleep(1)
		SendKeys.SendKeys('{TAB}')
		SendKeys.SendKeys('^{a}') #select all
		time.sleep(1)
		SendKeys.SendKeys('^{c}') #copy
		SendKeys.SendKeys('^{w}') #close page source window
	
	else:
		x = win32gui.FindWindow(None, "Certificate Management Center - Microsoft Internet Explorer")
		win32gui.SetForegroundWindow(x) #bring the window with x hwnd to the front
		s = win32gui.GetForegroundWindow()
		while x == s:
			SendKeys.SendKeys('%v')
			time.sleep(.5)
			SendKeys.SendKeys('c')
			time.sleep(.5)
			s = win32gui.GetForegroundWindow()
		SendKeys.SendKeys('^{a}') #select all
		SendKeys.SendKeys('^{c}') #copy	
		SendKeys.SendKeys('%{F4}') #MAKE SURE THE SOURCE WINDOW IS CURRENTLY ON TOP
			
elif re.search('.*Firefox.*', stuff):
	#print "Firefox"
	x = win32gui.FindWindow(None, "Certificate Management Center - Mozilla Firefox") #get hwnd given the window title
	j = win32gui.FindWindow(None, "view-source: - Source of: identrust.com - Mozilla Firefox")
	win32gui.SetForegroundWindow(x) #bring the window with x hwnd to the front
	SendKeys.SendKeys('^{u}') #view page source in firefox
	#time.sleep(1)
	for p in range(5):
		if j != 0:
			break
		else:
			time.sleep(.5)
			j = win32gui.FindWindow(None, "view-source: - Source of: IdenTrust.com- Mozilla Firefox")
	SendKeys.SendKeys('{TAB}')
	SendKeys.SendKeys('^{a}') #select all
	time.sleep(1)
	SendKeys.SendKeys('^{c}') #copy
	SendKeys.SendKeys('^{w}') #close page source window	
elif re.search('.*Internet Explorer.*', stuff):
	#print "Internet Explorer"	
	x = win32gui.FindWindow(None, "Certificate Management Center - Microsoft Internet Explorer")
	win32gui.SetForegroundWindow(x) #bring the window with x hwnd to the front
	s = win32gui.GetForegroundWindow()
	while x == s:
		SendKeys.SendKeys('%v')
		time.sleep(.5)
		SendKeys.SendKeys('c')
		time.sleep(.5)
		s = win32gui.GetForegroundWindow()
	SendKeys.SendKeys('^{a}') #select all
	SendKeys.SendKeys('^{c}') #copy	
	SendKeys.SendKeys('%{F4}') #MAKE SURE THE SOURCE WINDOW IS CURRENTLY ON TOP
win32clipboard.OpenClipboard(0)
junk = win32clipboard.GetClipboardData(win32con.CF_TEXT)
win32clipboard.EmptyClipboard()
win32clipboard.CloseClipboard() #IMPORTANT must close before using again (?)
####### find and replace extra spaces, Jr/Sr/Mr/Mrs/Ms, I/II/III/etc, "," "."
acct = re.findall('\d{8,}-\d+', junk)
fullx = re.findall('\>NAME\s.+\n.+(?P<name>\;[a-zA-Z -.]+)', junk)
#print fullx[0]
full = re.sub('[;.,]', '', fullx[0])
full = re.compile('[mjs]r', flags=2).sub('', full) #how to ignore case
#full = re.sub('\W+I{2,}', '', full) #get rid of I, II, III
full = re.sub('(\W+I{2,})|(\W+IV)', '', full)
full = string.rstrip(string.lstrip(full))
#full = full.title() # only do title on names w/ all caps or no caps
if full.islower() | full.isupper():
	full = full.title()
#emailx = re.findall('E-MAIL.+\n.+(?P<mail>\;.+@.+\.\w{2,})', junk)
emailx = re.findall(';.+@.+\.\w{2,}', junk)
email = re.sub('\;', '', emailx[0])

phone = re.findall('\d{3}\W\d{3}\W\d{4}', junk)

n = full.split(' ')
last = n[-1]
login = full[0]+n[-1]+'-'+acct[0]
win32clipboard.OpenClipboard(0)
win32clipboard.SetClipboardText(acct[0])

for row in topWindows:
	y = re.search('Remedy', row[1], 2)
	if y:
		myWindow = [row[0],row[1],win32gui.GetClassName(row[0])]	
		y = win32gui.FindWindow(None, myWindow[1]) #get hwnd of Remedy
		win32gui.SetForegroundWindow(y) #bring Remedy window to front
ar = win32com.client.Dispatch("Remedy.User")
x = ar.OpenForm(0, 'IdenTrust.com', 'HPD:HelpDesk', 1, 1)
x.GetField('Account #').Value = acct[0]
x.GiveFieldFocus('Status')
SendKeys.SendKeys('r')
x.GiveFieldFocus('Account #')
SendKeys.SendKeys('{ENTER}')  #search off acct num
    #if nothing found off acct num, look for error window
time.sleep(1)
e = win32gui.FindWindow(None, "Remedy User - Error")
time.sleep(.25)
if e > 0:
	SendKeys.SendKeys('{ENTER}') #close error window	
	#x.GetField('Account #').Value = '' #clear out the acct num field	
	p = ar.OpenForm(0, 'identrust.com', 'SHR:People', 1, 1) #to create new profile
	p.GetField('Login Name').Value = login
	p.GetField('Account #').Value = acct[0]
	p.GetField('Last Name').Value = last
	p.GetField('First Name').Value = n[0]
	p.GetField('Full Name').Value = full
	p.GetField('Email').Value = email
	p.GetField('Phone').Value = phone[0]
	x.GiveFieldFocus('Account #')
	#user should review profile data before saving	
win32clipboard.EmptyClipboard()
win32clipboard.CloseClipboard()
print e
print login
print acct[0]
print full
print email
print phone[0]