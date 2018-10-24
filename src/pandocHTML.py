#TODO:
# 1. strip id from figure legend text
# 2. assemble links in side menu rather than make them non-functional

import pypandoc
import os
import sys
import re
from shutil import copyfile

# tell browser we are using UTF-8 otherwise many unicode characters are handled incorrectly
# provide link to external style sheet
header = '''\
<!DOCTYPE html>
<meta charset="UTF-8">
<link href="normalstyle.css" rel="stylesheet" type="text/css">
<html>
<head >
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js">
</script>

<script>
$(function(){
  $('img').click(function (e) {
    if (e.shiftKey){
      $(this).animate({
         height: $(this).height()/1.25,
         width: $(this).width()/1.25
      });
    } else {
      $(this).animate({
          height: $(this).height()*1.25,
          width: $(this).width()*1.25
      });
    }
  });
});
</script>

</head>
<body>

'''

sideMenu = '''\
<ul>
  <li><a href= '{0}' >&#8679;</a></li>
  <li><a href= '{1}' >&#8678;</a></li>
  <li><a href= '{2}' >&#8680;</a></li>
</ul>
<div class='main_text'>

'''
footer = '''
</div class='main_text'> 
</body>
</html>
'''


def cleanHTML(filePath, fileDirectory):
    html = pypandoc.convert_file(filePath, 'html5', extra_args=['--extract-media='+fileDirectory])
    # use local reference to media folder
    html = re.sub(r'img.*?media', r'img src="media', html)
    # remove figure legend id which is too long
    html = re.sub(r' id="figure.*?>', r'> ', html)
    # change h5 tags to class p.legend
    html = html.replace('<h5', '<p class="legend"')
    html = html.replace("</h5>", "</p>")
    # get rid of paragraph marks in lists so that list CSS formating works
    html = html.replace('<li><p>', '<li>')
    html = html.replace('</li><p>', '</li>')
    # create container for table so that it can be centered to the same width as the text and images 
    html = html.replace('<table>', '<div class="tablecontainer">\n<table>')
    html = html.replace('</table>', '</table>\n</div>')
    return html

def subChapters(html, fileName, fileDirectory, sideMenu):
    homeFileName = fileName + '.html'
    
    h2str = '<h2 id="' 
    pages = html.split(h2str)   # split file at h2 mark to get subsections
    lastFileName = homeFileName
    index = ''
    # ignore first subsection which will contain the index
    for i in range(1, len(pages)):
        currentFileName = fileName + '_' + str(i) + '.html'
        nextFileName = fileName + '_' + str(i+1) + '.html'
        menu = sideMenu.format(homeFileName, lastFileName, nextFileName)
        
        # create index
        h2line = pages[i].split('\n', 1)[0] # get first line, which has title
        subtitle = h2line.split('">')[1]
        subtitle = subtitle.split('</h2>')[0]
        link = '<h4><a href= "{0}" > {1} </a></h4>'
        link = link.format(currentFileName, subtitle)
        index = index + link + '\n'
        
        # remove forward icon from last page
        if i == len(pages)-1:
            menu = menu.replace('&#8680;', '') 

        html = header + menu + h2str + pages[i] + footer  # add h2str back after split
        outFile = os.path.join(fileDirectory, currentFileName)
        with open(outFile, 'w', encoding="utf-8") as f:
            f.write(html)
        lastFileName = currentFileName
        
    # create first page
    firstPage = pages[0]
    firstLine = firstPage.split('\n', 1)[0] + '\n' # get first line
    everythingelse = firstPage.split('\n', 1)[1] 
    menu = menu.replace('&#8678;', '') # remove forward icon from first page
    menu = menu.replace('&#8680;', '') # remove backward icon from first page
    html = header + menu + firstLine + index + everythingelse + footer 
    outFile = os.path.join(fileDirectory, fileName + '.html')
    with open(outFile, 'w', encoding="utf-8") as f:
            f.write(html)

def bigFile(html, fileName, fileDirectory):
    outFile = os.path.join(fileDirectory, fileName + '_Big.html' )  
    with open(outFile, 'w', encoding="utf-8") as f:
        f.write(html)
        
def copyStyleSheet(htmlDirectory):
    src = r'stylesheets/normalstyle.css'
    dst = os.path.join(htmlDirectory, 'normalstyle.css')
    copyfile(src, dst)


if len(sys.argv) == 2:
    filePath = sys.argv[1]
else:
    filePath = r'/Users/david/Documents/Text Book/text/C1.docx'
#     filePath = r'C:\Users\dmckinnon\Documents\9 Software\Python\Doc2HTM\src\text\C1.docx'  # windows

baseName = os.path.basename(filePath)
fileDirectory = os.path.dirname(filePath)
fileName, fileExtension = os.path.splitext(baseName)
htmlDirectory = os.path.join(fileDirectory, fileName + '_html')

  
print('filePath = ', filePath)                                                                                                                                                                                                                                                            
print('fileName = ', fileName)
print('fileDirectory = ', fileDirectory)
print('htmlDirectory = ', htmlDirectory)


# get a cleaned up HTML string from the word document
html = cleanHTML(filePath, htmlDirectory)  
# print(html)
# Note: pandoc creates the htmlDirectory in order to store the media files
subChapters(html, fileName, htmlDirectory, sideMenu)
bigFile(html, fileName, htmlDirectory)
copyStyleSheet(htmlDirectory)




