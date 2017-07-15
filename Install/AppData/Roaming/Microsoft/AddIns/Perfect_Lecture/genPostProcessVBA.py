import sys
import os
import shutil
import time
import xml.etree.ElementTree as et
import logging

# Logging

#logging.basicConfig(stream=sys.stderr, level=logging.INFO)  # RELEASE
logging.basicConfig(stream=sys.stderr, level=logging.DEBUG)  # DEBUG



def pause_exit():
    ## DEBUG
    os.system('pause') #DEBUG
    exit()

def quotedStr(str_in):
    return '"' + str_in + '"'

def writeNotePage(post_process_script_file, pageNum, lines_in_currentPage):
    print('writeNotePage,{},{}'.format(pageNum, quotedStr('<script>')), file=post_process_script_file)
    for tmp_line in lines_in_currentPage.splitlines(): #TODO: simplify this
        print('addNewLineToNotePage,{},{}'.format(pageNum, quotedStr(tmp_line)), file=post_process_script_file)
    print('addNewLineToNotePage,{},{}'.format(pageNum, quotedStr('</script>')), file=post_process_script_file)

# Program Setup
print("genPostProecessVBA.py")

workDir = os.path.dirname(os.path.realpath(__file__))
codecDir = os.path.join(workDir, 'ImageMagick-portable')  #PATH
ttsDir = os.path.join(workDir, 'TTS_engine')  #PATH
tmpDir = 'C:\\Temp'                                         #PATH

if not os.path.exists(tmpDir):
    os.makedirs(tmpDir)

# Get work dir & filename of PDF file
# [TODO] check sys.argv[1] isn't blank
script_file = sys.argv[1] + '.script.xml'
pdf = sys.argv[1] + '.pdf'
pdfRoot = os.path.basename(sys.argv[1])
pdfDir = os.path.dirname(pdf)

default_framerate = 1/5  # default_framerate = 1/5 frames per second = 12 frames per minute
MAX_FRAME_RATE = 10**5   # [TODO]: adjust

numPageAdded = 0

# Step1: Parse Perfect Lecture Script & Run TTS
post_process_script  = os.path.join(tmpDir, 'post_process.iscript')  #PATH
with open(post_process_script, 'w', encoding = 'UTF-8') as post_process_script_file:
    tree = et.ElementTree(file=script_file)
    root = tree.getroot()
    for child in root:

        logging.debug('%s %s %s %s', child.tag, child.attrib, child.text, child.tail) #DEBUG
        pageNum = int(child.attrib['index']) + numPageAdded #DEBUG
        # for grand_child in child: #DEBUG
        #     logging.debug('%s %s %s %s', grand_child.tag, grand_child.attrib, grand_child.text, grand_child.tail) #DEBUG

        #[DONE]: more than 1 script section?
        script_text = ''
        for script in child.findall('script'):
            script_text += script.text
            logging.debug('%s %s %s %s', script.tag, script.attrib, script.text, script.tail) #DEBUG

        if script_text != '':
            #parse script
            logging.debug('%s', 'Enter if script_text') #DEBUG
            # [TODO]: parse Perfect Lecture Script from script.text

            # implement C-style multi-line comments:  /* ... */
            in_js_comment  = False
            processed_text = ''
            remain_text = script_text
            while remain_text != '':
                if in_js_comment:
                    if '*/' in remain_text:
                        split_text = remain_text.split('*/', maxsplit=1)
                        remain_text = split_text[-1]
                        in_js_comment = False
                    else:
                        logging.error('   [ERROR] Find comments starting with dangling /*: the corresponding */ is not found!')
                        logging.error('   [NOTE]  Script after removing /* ... */ comments: \n%s', processed_text+remain_text)
                        pause_exit()
                else:
                    if '/*' in remain_text:
                        split_text = remain_text.split('/*', maxsplit=1)
                        processed_text += split_text[0]
                        remain_text = split_text[-1]
                        in_js_comment = True
                    else:
                        processed_text += remain_text
                        remain_text = ''
                        break

            if '*/' in processed_text:
                logging.error('   [ERROR] Find comments ending with dangling */: the corresponding /* is not found!!')
                logging.error('   [NOTE]  Script after removing /* ... */ comments: \n%s', processed_text)
                pause_exit()
            logging.debug('%s', processed_text) #DEBUG

            # BEGIN DUPLICATE PAGES
            #[TODO]: implement lines_in_currentPage
            lines_in_currentPage = ''
            # END DUPLICATE PAGES
            for line in processed_text.splitlines():
                processed_line = line
                # implement XML-style comments:  <!-- ... -->  by  et.ElementTree
                # implement C-style single line comments:  // ...
                processed_line = processed_line.split('//', maxsplit=1) [0].strip().replace('“', '"').replace('”', '"').replace('‘', '"').replace("’", '"').replace("`", '"').replace("'", '"') #TODO: hot fix (ad-hoc)

                # logging.debug('   processed_line = \'%s\'', processed_line) #DEBUG
                # cmd_opt = processed_line.split()
                # logging.debug('   cmd_opt = \'%s\'', str(cmd_opt)) #DEBUG

                # [TODO]: CHECK IT (BEGIN)
                # [DONE]: write customized tokenizer: deal with other white space characters (e.g. \t) enclosed by ""
                # [TODO]: deal with only the "" beside the separator
                logging.debug('   processed_line = \'%s\'', processed_line) #DEBUG
                
                cmd_opt = []

                remain_text = processed_line
                InField = False
                MergeWithPrev = False
                field = ''
                while remain_text != '':
                    if InField:
                        try:
                            pos = remain_text.index('"""')
                        except:
                            cmd_opt.append(field + remain_text)
                            remain_text = ''

                            logging.debug('Infield   field = \'%s\'', field) #DEBUG
                            logging.debug('Infield   remain_text = \'%s\'', remain_text) #DEBUG
                        else:
                            field += remain_text[:pos] # Remove """
                            cmd_opt.append(field)
                            remain_text = remain_text[pos+3:]

                            logging.debug('Infield   field = \'%s\'', field) #DEBUG
                            logging.debug('Infield   remain_text = \'%s\'', remain_text) #DEBUG
                        field = ''
                        InField = False
                    else:
                        try:
                            pos = remain_text.index('"""')
                        except:
                            cmd_opt += remain_text.strip().split()

                            remain_text = ''
                            InField = False

                            logging.debug('Not Infield   field = \'%s\'', field) #DEBUG
                            logging.debug('Not Infield   remain_text = \'%s\'', remain_text) #DEBUG
                        else:
                            cmd_opt += remain_text[:pos].strip().split() # Remove """
                            
                            remain_text = remain_text[pos+3:]
                            InField = True

                            logging.debug('Not Infield   field = \'%s\'', field) #DEBUG
                            logging.debug('Not Infield   remain_text = \'%s\'', remain_text) #DEBUG                        
                        field = ''
                
                if field != '':
                    cmd_opt.append(field)

                for j in range(len(cmd_opt)):
                    if cmd_opt[j][:3] == '"""':
                        cmd_opt[j] = cmd_opt[j][3:]                        
                    if cmd_opt[j][-3:] == '"""':
                        cmd_opt[j] = cmd_opt[j][:-3]

                logging.debug('   cmd_opt = \'%s\'\n\n', str(cmd_opt)) #DEBUG

                # [TODO]: CHECK IT (END)

                if len(cmd_opt) >= 2:
                    cmd = cmd_opt[0].lower()
                    opt = cmd_opt[1:]
                    logging.debug('   (cmd, opt) = (%s, %s)', cmd, opt) #DEBUG                        
                    
                    # BEGIN DUPLICATE PAGES

                    # [TODO]:複製該頁，pageNum += 1
                    if cmd == 'transpose' or cmd == 'point':
                        print('duplicate_page,{}'.format(pageNum), file=post_process_script_file)
                        
                        # [TODO]:切出前半的放到目前頁，切出後半的note放進下一頁
                        writeNotePage(post_process_script_file, pageNum, lines_in_currentPage)
                        lines_in_currentPage = processed_line + '\n' #TODO: or + '\r\n' ?
                        pageNum += 1
                        numPageAdded += 1
                    else:                        
                        #[TODO]: implement lines_in_currentPage
                        lines_in_currentPage += processed_line + '\n' #TODO: or + '\r\n' ?

                    # END DUPLICATE PAGES

                    if cmd == 'transpose':
                        if len(opt) < 3:
                            logging.error('   [ERROR] The number of arguments of "transpose" command is less than 3.')
                            pause_exit()

                        # [TODO]: 和TTS結合？
                        # [DONE]產生修改的LaTeX碼 （用目前頁碼）
                        # [TODO]: implement replace, highlight, cancel commands

                        transpose_obj = opt[0]
                        transpose_cmd = opt[1].lower()

                        if transpose_cmd == 'replace':
                            if len(opt) != 4:
                                logging.error('   [ERROR] Usage: transpose <Obj> replace <Find> <Replacement>')
                                pause_exit()
                            find = opt[2]
                            replacement = opt[3]

                            # [TODO]: NOW
                            # 寫vba碼
                            # Step1: 切換到對應頁 
                            # Step2: replace obj, find, replacement #TODO: 需要用str.encode() & str.decode()處理編碼或是直接貼上？
                            print('edit_equation,{},{},{},{}'.format(pageNum, quotedStr(transpose_obj), quotedStr(find), quotedStr(replacement)), file=post_process_script_file)
                    elif cmd == 'point':
                        if len(opt) != 4 and len(opt) != 6:
                            logging.error('   [ERROR] Usage: point <Position_X_Ratio> <Position_Y_Ratio> <Pointer_Type> <Pointer_Color> [<pointer_width> <pointer_height>].')
                            pause_exit()
                        else:
                            pointer_posX = opt[0]
                            pointer_posY = opt[1]
                            pointer_type = opt[2].lower()
                            pointer_color = opt[3].lower()

                            print(pointer_type)
                            
                            # [TODO]: Implement other colors
                            if pointer_color == 'black':
                                pointer_R = 0
                                pointer_G = 0
                                pointer_B = 0
                            elif pointer_color == 'white':
                                pointer_R = 255
                                pointer_G = 255
                                pointer_B = 255
                            elif pointer_color == 'red':
                                pointer_R = 255
                                pointer_G = 0
                                pointer_B = 0
                            elif pointer_color == 'orange':
                                pointer_R = 255
                                pointer_G = 165
                                pointer_B = 0
                            elif pointer_color == 'darkorange':
                                pointer_R = 255
                                pointer_G = 140
                                pointer_B = 0
                            elif pointer_color == 'yellow':
                                pointer_R = 255
                                pointer_G = 255
                                pointer_B = 0
                            elif pointer_color == 'green':
                                pointer_R = 255
                                pointer_G = 0
                                pointer_B = 0
                            elif pointer_color == 'blue':
                                pointer_R = 0
                                pointer_G = 0
                                pointer_B = 255
                            elif pointer_color == 'magenta' or pointer_color == 'purple':
                                pointer_R = 255
                                pointer_G = 0
                                pointer_B = 255
                            elif pointer_color == 'cyan':
                                pointer_R = 0
                                pointer_G = 255
                                pointer_B = 255
                            else:
                                # Default: Red
                                pointer_R = 255
                                pointer_G = 0
                                pointer_B = 0
                            

                            if len(opt) == 6:
                                pointer_width = opt[4]
                                pointer_height = opt[5]
                            else:
                                if pointer_type == "arrow":
                                    pointer_width = 50
                                    pointer_height = 50
                                elif pointer_type == "circle":
                                    pointer_width = 5
                                    pointer_height = 5
                                else:
                                    pointer_width = 50
                                    pointer_height = 50

                            if pointer_type == "arrow":
                                pointer_rotation = 45
                            elif pointer_type == "circle":
                                pointer_rotation = 0
                            else:
                                pointer_rotation = 0

                            print('addPointer,{},{},{},{},{},{},{},{},{},{}'.format(pageNum, quotedStr(pointer_type), pointer_R, pointer_G, pointer_B, pointer_posX, pointer_posY, pointer_width, pointer_height, pointer_rotation), file=post_process_script_file)

            # BEGIN DUPLICATE PAGES
            #[TODO]: write out remaining lines in current note page
            writeNotePage(post_process_script_file, pageNum, lines_in_currentPage)
            lines_in_currentPage = '' #TODO: or + '\r\n' ?
            # END DUPLICATE PAGES


pause_exit()