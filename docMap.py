import sys
import win32com.client
from datetime import datetime, timedelta
from jinja2 import BaseLoader, Template
import codecs
import json
from collections import namedtuple

olFolderTodo = 28

outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")
todo_folder = ns.GetDefaultFolder(olFolderTodo)
todo_items = todo_folder.Items

CommandTasks = []
BugLists = []
OOTO = None
BlockingItems = []

class TaskList:
    subject = ''
    body = []
    flag = ''

class BugId:
    bugId = 0
    description = ''
    lastComment = ''
    count = 0

class OOTOMsg:
    date = ''
    reason = ''

class COST:
    def STRIKE():
        return "STRIKE"
    def RED():
        return "RED"

def print_encoded(s):
    print(s.encode(sys.stdout.encoding, 'replace'))

def PrintTODO():
    for i in range(1, 1 + len(todo_items)):
        item = todo_items[i]
        if item.__class__.__name__ == '_MailItem':
            print_encoded(u'Email: {0}. Due: {1}'.format(item.Subject, item.TaskDueDate))
        elif item.__class__.__name__ == '_ContactItem':
            print_encoded(u'Contact: {0}. Due: {1}'.format(item.FullName, item.TaskDueDate))
        elif item.__class__.__name__ == '_TaskItem':
            print_encoded(u'Task: {0}. Due: {1}'.format(item.Subject, item.DueDate))
        else:
            print_encoded(u'Unknown Item type: {0}'.format(item))

def printBugContent(item):
    sep = '------------------------------------------------------------------------------'
    body = item.Body
    s = 0
    e = 1

    first = 'has been reported with the following information:'
    last = 'Reply to this Comment'
    newc = 'NEW Comment: '
    postfix = ''
    try:
        s = body.index(first) + len(first) + len(sep)
        s = body.index(newc) + len(newc)
        e = body.index('__', s)
        if (e-s) > 400:
            e = 400
            postfix = ' ...'

        return body[s:e] + postfix
    except:
        pass
    return None

def SearchBugId(id):
    global BugLists
    index = 0
    for b in BugLists:
        if id == b.bugId:
            return index
        index = index + 1
    return -1

def IsBugId(item):
    b = '[B]'
    arb = 'ARB'

    if b in item.Subject and arb in item.Subject:
        return True
    return False

def GetBugId(str):
    b = '[B]'
    arb = 'ARB'
    ret = str
    s = ret.index(b) + len(b)
    e = ret.index(arb, s)
    return ret[s:e].strip()

def GetBugDescription(sub):
    arb = 'ARB'
    return sub[sub.index(arb)+len(arb):]

def ProcessBug(item):
    autoReply = 'Autometic Reply'
    if autoReply in item.Subject:
        return None
    global BugLists
    id = GetBugId(item.Subject)
    idx = SearchBugId(id)
    bugId = BugId()
    bugId.bugId = id
    bugId.description = GetBugDescription(item.Subject)
    bugId.lastComment = printBugContent(item)

    if idx == -1:
        BugLists.append(bugId)
    else:
        bugId.count = BugLists[idx].count + 1
        BugLists[idx] = bugId


def PrintBugs(item):
    if IsBugId(item) == True:
        print(item.Subject)
        print(printBugContent(item))

def printWeekTask(item):
    cur = datetime.now()
    dd = item.TaskDueDate.replace(tzinfo=None)
    cd = cur.replace(tzinfo=None)

    ws = cd - timedelta(days=cd.weekday())
    we = ws + timedelta(days=15)
    if dd > ws and dd < we:
        print_encoded(u'Task: {0}'.format(item.Subject))
        print(printItemContent(item))

def IsPending(item):
    cur = datetime.now()
    dd = item.TaskCompletedDate.replace(tzinfo=None)
    return cur < dd

CMD_SIG = 'PYBOT CMD:'

def IsOotoCmd(item):
    global CMD_SIG
    sig = CMD_SIG + " OOTO"
    return sig in item.Subject

def IsBlockingCmd(item):
    global CMD_SIG
    sig = CMD_SIG + " BLOCKING"
    return sig in item.Subject

def IsAccessCmd(item):
    sub = item.Subject
    global CMD_SIG
    if CMD_SIG in sub and not IsOotoCmd(item) and not IsBlockingCmd(item):
        return True
    return False

def GetBlockingData(str):
    bullet = '*'
    body = []
    htmlTagStart = ''
    htmlTagEnd = ''
    bstr = str
    index = 0
    bHtmlInit = False
    while True:
        try:
            b = bstr.index('*')+1
            if '*' in bstr[b:]:
                e = bstr.index('*', b)
            else:
                e = len(bstr) - 1
            newStr = bstr[b:e]
            if 'From' in newStr:
                bHtmlInit = True
                newStr = newStr[0:newStr.index('From')]
            body.append(htmlTagStart + newStr + htmlTagEnd)
            bstr = bstr[e:]
            index = index + 1
            if bHtmlInit == True:
                htmlTagStart = '&nbsp;&nbsp;<STRIKE>'
                htmlTagEnd = '</STRIKE>'
                bHtmlInit = False
        except:
            if index == 0:
                body.append(str)
            return body
    return None

def ProcessBlocking(item):
    body = item.Body
    global BlockingItems
    BlockingItems.extend(GetBlockingData(body))

def ProcessOOTO(item):
    global OOTO
    content = item.Body
    print(content)
    OOTO = OOTOMsg()
    s = content.index('*')+1
    e = content.index('*', s)
    OOTO.date = content[s:e]
    content = content[e:]
    OOTO.reason = content[content.index('*')+1:]

def MarkCompleted(item):
    item.TaskCompletedDate = datetime.now() - timedelta(days=1)
    item.Save()


def SearchInListSub(sub):
    global CommandTasks
    idx = 0
    for o in CommandTasks:
        if o.subject == sub:
            return idx
        idx = idx + 1
    return -1

def RenderFile():
    templateFile = 'status.html'
    outFile = 'newStatus.html'
    f = codecs.open(templateFile, 'r')
    source = f.read()
    t = Template(source)
    r = t.render(CommandTasks=CommandTasks, BugLists=BugLists, OOTO=OOTO, BlockingItems=BlockingItems)

    newFile = open(outFile, 'w', encoding="utf-8")
    newFile.write(r)
    newFile.close()
    return r

def UpdateBody(str):
    body = []
    bstr = str
    index = 0
    try:
        bstr = bstr[0:bstr.index('From')]
    except:
        pass
    while True:
        try:
            b = bstr.index('*')+1
            if '*' in bstr[b:]:
                e = bstr.index('*', b)
            else:
                e = len(bstr) - 1
            body.append(bstr[b:e])
            bstr = bstr[e:]
            index = index + 1
        except:
            if index == 0:
                body.append(str)
            return body
    return None

def ProcessCommandOrg(item):
    Subject = item.Subject
    global CMD_SIG
    global CommandTasks
    s = Subject.index(CMD_SIG) + len(CMD_SIG)
    sub = Subject[s:]
    idx = SearchInListSub(sub)

    st = TaskList()
    st.subject = sub
    st.body = UpdateBody(item.Body)

    if idx == -1:
        CommandTasks.append(st)
    else:
        # print("Found : ", st.subject)
        CommandTasks[idx] = st

def IsAccessCmdRE(item):
    global CMD_SIG
    reCmd = "RE: " + CMD_SIG
    if reCmd in item.Subject:
        return True
    return False

def UpdateBodyRE(str):
    body = []
    bstr = str
    index = 0
    htmlTagStart = ''
    htmlTagEnd = ''
    bHtmlInit = False
    while True:
        try:
            b = bstr.index('*')+1
            if '*' in bstr[b:]:
                e = bstr.index('*', b)
            else:
                e = len(bstr) - 1
            newStr = bstr[b:e]
            if 'From' in newStr:
                bHtmlInit = True
                newStr = newStr[0:newStr.index('From')]
            body.append(htmlTagStart + newStr + htmlTagEnd)
            bstr = bstr[e:]
            index = index + 1
            if bHtmlInit == True:
                htmlTagStart = '&nbsp;&nbsp;&nbsp;&nbsp<STRIKE>'
                htmlTagEnd = '</STRIKE>'
                bHtmlInit = False
        except:
            if index == 0:
                body.append(str)
            return body
    return None

def ProcessCommandRE(item):
    Subject = item.Subject
    global CMD_SIG
    global CommandTasks
    s = Subject.index(CMD_SIG) + len(CMD_SIG)
    sub = Subject[s:]
    idx = SearchInListSub(sub)

    st = TaskList()
    st.subject = sub
    st.body = UpdateBodyRE(item.Body)

    if idx == -1:
        CommandTasks.append(st)
    else:
        CommandTasks[idx] = st

def ProcessDueTaskBody(item):
    body = item.Body
    # print("Porcessing for ", item.Subject)
    try:
        frm = 'From'
        body = body[0:body.index(frm)]
    except:
        pass

    keywords = ['Thanks', 'nvcr', 'review', 'NVCR']
    if any(k in body for k in keywords):
        return "Under Review"
    return body

def ProcessDueTask(item):
    global CommandTasks
    subject = item.Subject
    body = item.Body
    try:
        rep = 'RE: '
        subject = subject[subject.index(rep) + len(rep):]
    except:
        pass

    st = TaskList()
    st.subject = subject
    b = ProcessDueTaskBody(item)
    st.body = []
    st.body.append(b)
    CommandTasks.append(st)

def ProcessItem(item):
    if IsPending(item) == False:
        return
    if IsAccessCmdRE(item):
        ProcessCommandRE(item)
    elif IsAccessCmd(item):
        ProcessCommandOrg(item)
    elif IsOotoCmd(item):
        ProcessOOTO(item)
    elif IsBlockingCmd(item):
        ProcessBlocking(item)
    elif IsBugId(item):
        ProcessBug(item)
    else:
        ProcessDueTask(item)
        #MarkCompleted(item)

for i in range(1, 1 + len(todo_items)):
    item = todo_items[i]
    if item.__class__.__name__ == '_MailItem':
        ProcessItem(item)
# RenderFile()

def GetStatus():
    global CommandTasks
    global BugLists
    global OOTO
    global BlockingItems

    CommandTasks = []
    BugLists = []
    OOTO = None
    BlockingItems = []

    for i in range(1, 1 + len(todo_items)):
        item = todo_items[i]
        if item.__class__.__name__ == '_MailItem':
            ProcessItem(item)
    return RenderFile()
