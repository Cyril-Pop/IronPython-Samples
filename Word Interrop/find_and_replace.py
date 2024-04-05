# Code updated for Net Core 8 compatibility (with IronPython 2 or 3)
import sys
import clr
import System
import time

clr.AddReference("System.Runtime.InteropServices")
import System.Runtime.InteropServices


class WordlEnum:
    """
    values Enum from API Interop Mircrosoft Doc , to avoid import lib in clr for Enum 
    """
    Word_WdReplace_wdReplaceAll = 2
    Word_WdReplace_wdReplaceNone = 0
    Word_WdReplace_wdReplaceOne = 1
    Word_WdFindWrap_wdFindContinue = 1

def find_replace(objRng, search_txt, replce_txt):
    # missing = System.Type.Missing
    missing = None
    if len(replce_txt) < 250:
        objRng.Find.Execute(
        search_txt, # search text
        True, # match case
        missing, # match whole word
        missing, # match wildcards
        missing, # match sounds like
        missing, # match all word forms
        True, # forward?
        WordlEnum.Word_WdFindWrap_wdFindContinue, # wrap enum
        missing, # format?
        replce_txt, # replace with
        WordlEnum.Word_WdReplace_wdReplaceAll, # replace enum
        missing, # match Kashida (arabic letters)
        missing, # diatrics?
        missing, # match other Arabic letters
        missing # support right to left
        )
    else:
        i =0
        flag = True
        while flag and i < 20: # prevent infiniteloop
            copyRange = objRng.Duplicate 
            flag = copyRange.Find.Execute(
            search_txt, # search text)
            missing, # match case
            missing, # match whole word
            missing, # match wildcards
            missing, # match sounds like
            missing, # match all word forms
            missing, # forward?
            missing, # wrap enum
            missing, # format?
            missing, # replace with
            missing, # replace enum
            missing, # match Kashida (arabic letters)
            missing, # diatrics?
            missing, # match other Arabic letters
            missing # support right to left
            )
            if not flag:
                break
            else:
                copyRange.Text = replce_txt
            i += 1
      
def doc_replace_text(source_filename, tokens, values):
    global errors
    # word_application = Word.ApplicationClass()
    word_application = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Word.Application", True))
    word_application.visible = True
    document = word_application.Documents.Open(source_filename)
    #time.sleep(2)
    #Find and Replace Process
    for _find, _replace in zip(tokens,values):
        for myStoryRange  in document.StoryRanges:
            find_replace(myStoryRange , _find, _replace)
            try:
                while myStoryRange.NextStoryRange is not  None:
                    q = myStoryRange.NextStoryRange 
                    find_replace(q, _find, _replace)
            except:
                import traceback
                errors.append(traceback.format_exc())
            #Find and replace in TextBox(shapes)
            try:    
                for shape in document.Shapes:
                    initialText = shape.TextFrame
                    if initialText.HasText:
                        rangeobj = initialText.TextRange 
                        find_replace(rangeobj, _find, _replace)

            except:
                import traceback
                errors.append(traceback.format_exc())
                                
toList = lambda x : x if hasattr(x, '__iter__') else [x]
lst_filenames = toList(IN[0])
tokens = IN[1]
values = IN[2]
errors = []    
word_application = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Word.Application", True))
word_application.Quit()
word_application = None
for filename in lst_filenames:
    doc_replace_text(filename, tokens, values) 

OUT = errors
import sys
import clr
import System
import time

clr.AddReference("System.Runtime.InteropServices")
import System.Runtime.InteropServices


class WordlEnum:
    """
    values Enum from API Interop Mircrosoft Doc , to avoid import lib in clr for Enum 
    """
    Word_WdReplace_wdReplaceAll = 2
    Word_WdReplace_wdReplaceNone = 0
    Word_WdReplace_wdReplaceOne = 1
    Word_WdFindWrap_wdFindContinue = 1

def find_replace(objRng, search_txt, replce_txt):
    # missing = System.Type.Missing
    missing = None
    if len(replce_txt) < 250:
        objRng.Find.Execute(
        search_txt, # search text
        True, # match case
        missing, # match whole word
        missing, # match wildcards
        missing, # match sounds like
        missing, # match all word forms
        True, # forward?
        WordlEnum.Word_WdFindWrap_wdFindContinue, # wrap enum
        missing, # format?
        replce_txt, # replace with
        WordlEnum.Word_WdReplace_wdReplaceAll, # replace enum
        missing, # match Kashida (arabic letters)
        missing, # diatrics?
        missing, # match other Arabic letters
        missing # support right to left
        )
    else:
        i =0
        flag = True
        while flag and i < 20: # prevent infiniteloop
            copyRange = objRng.Duplicate 
            flag = copyRange.Find.Execute(
            search_txt, # search text)
            missing, # match case
            missing, # match whole word
            missing, # match wildcards
            missing, # match sounds like
            missing, # match all word forms
            missing, # forward?
            missing, # wrap enum
            missing, # format?
            missing, # replace with
            missing, # replace enum
            missing, # match Kashida (arabic letters)
            missing, # diatrics?
            missing, # match other Arabic letters
            missing # support right to left
            )
            if not flag:
                break
            else:
                copyRange.Text = replce_txt
            i += 1
      
def doc_replace_text(source_filename, tokens, values):
    global errors
    # word_application = Word.ApplicationClass()
    word_application = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Word.Application", True))
    word_application.visible = True
    document = word_application.Documents.Open(source_filename)
    #time.sleep(2)
    #Find and Replace Process
    for _find, _replace in zip(tokens,values):
        for myStoryRange  in document.StoryRanges:
            find_replace(myStoryRange , _find, _replace)
            try:
                while myStoryRange.NextStoryRange is not  None:
                    q = myStoryRange.NextStoryRange 
                    find_replace(q, _find, _replace)
            except:
                import traceback
                errors.append(traceback.format_exc())
            #Find and replace in TextBox(shapes)
            try:    
                for shape in document.Shapes:
                    initialText = shape.TextFrame
                    if initialText.HasText:
                        rangeobj = initialText.TextRange 
                        find_replace(rangeobj, _find, _replace)

            except:
                import traceback
                errors.append(traceback.format_exc())
                                
toList = lambda x : x if hasattr(x, '__iter__') else [x]
lst_filenames = toList(IN[0])
tokens = IN[1]
values = IN[2]
errors = []    
word_application = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Word.Application", True))
word_application.Quit()
word_application = None
for filename in lst_filenames:
    doc_replace_text(filename, tokens, values) 

OUT = errors
