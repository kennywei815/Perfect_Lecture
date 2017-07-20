import sys
import os
import shutil
import time
import xml.etree.ElementTree as et
import logging

# Logging

logging.basicConfig(stream=sys.stderr, level=logging.INFO)  # RELEASE
# logging.basicConfig(stream=sys.stderr, level=logging.DEBUG)  # DEBUG



def pause_exit():
    ## DEBUG
    os.system('pause') #DEBUG
    exit()

def quotedStr(str_in):
    return '"' + str_in + '"'

# Program Setup
print("Executing pdf2mp4.py...")

workDir = os.path.dirname(os.path.realpath(__file__))
codecDir = os.path.join(workDir, 'ImageMagick-portable')  #PATH
ttsDir = os.path.join(workDir, 'TTS_engine')              #PATH
tmpDir = 'C:\\Temp'                                       #PATH

if not os.path.exists(tmpDir):
    os.makedirs(tmpDir)

# Get work dir & filename of PDF file
# [TODO] check sys.argv[1] isn't blank
script_file = sys.argv[1] + '.script.xml'
pdf = sys.argv[1] + '.pdf'
pdfRoot = os.path.basename(sys.argv[1])
pdfDir = os.path.dirname(pdf)

sizeSpec = sys.argv[2]

#audio = os.path.join(tmpDir, pdfRoot + '.wav')
audio = os.path.join(tmpDir, 'source.wav')
image = os.path.join(tmpDir, pdfRoot + '.jpeg')
video = os.path.join(tmpDir, pdfRoot + '.mp4')
audio_specifier = os.path.join(tmpDir, pdfRoot + '-{}.wav')
image_specifier = os.path.join(tmpDir, pdfRoot + '-{}.jpeg')
video_specifier = os.path.join(tmpDir, pdfRoot + '-{}.mp4')
mp4list = os.path.join(tmpDir, pdfRoot + '.mp4list.txt')

text_specifier = os.path.join(tmpDir, pdfRoot + '-{}.xml')


numPage = 0
pageAudio = []
frameRate = []
default_frameRate = 1/5  # default_frameRate = 1/5 frames per second = 12 frames per minute
MAX_FRAME_RATE = 10**5   # [TODO]: adjust

# Step1: Parse Perfect Lecture Script & Run TTS
post_process_script  = os.path.join(tmpDir, 'post_process.iscript')  #PATH
with open(post_process_script, 'w', encoding = 'UTF-8') as post_process_script_file:
    tree = et.ElementTree(file=script_file)
    root = tree.getroot()
    for page in root:

        logging.debug('%s %s %s %s', page.tag, page.attrib, page.text, page.tail) #DEBUG
        # for grand_child in page:
        #     logging.debug('%s %s %s %s', grand_child.tag, grand_child.attrib, grand_child.text, grand_child.tail) #DEBUG
        pageNum_source = int(page.attrib['index'])
        pageNum = pageNum_source
        print('--------------------------------------------------------------------------------------')
        print('Processing page {}...'.format(pageNum_source))

        pageAudio.append(None)
        frameRate.append(default_frameRate)

        #[DONE]: more than 1 script section?
        script_text = ''
        for script in page.findall('script'):
            script_text += script.text
            logging.debug('%s %s %s %s', script.tag, script.attrib, script.text, script.tail) #DEBUG

        if script_text != '':
            #parse script
            logging.debug('%s', 'Enter if script_text') #DEBUG

            #PATH
            tts_text  = os.path.join(tmpDir, 'source.xml')              #PATH
            tts_audio = os.path.join(tmpDir, 'source.wav')              #PATH
            cur_tts_text  = text_specifier.format(numPage)
            cur_audio = audio_specifier.format(numPage)
            #cur_audio = audio
            pageAudio[-1] = cur_audio
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

            needTTS = False
            with open(cur_tts_text, 'w', encoding = 'UTF-8') as tts_text_file:
                for line in processed_text.splitlines():
                    processed_line = line
                    # implement XML-style comments:  <!-- ... -->  by  et.ElementTree
                    # implement C-style single line comments:  // ...
                    processed_line = processed_line.split('//', maxsplit=1) [0].strip().replace('“', '"').replace('”', '"').replace('‘', "'").replace("’", "'").replace("`", "'") #TODO: hot fix (ad-hoc)

                    # logging.debug('   processed_line = \'%s\'', processed_line) #DEBUG
                    # cmd_opt = processed_line.split()
                    # logging.debug('   cmd_opt = \'%s\'', str(cmd_opt)) #DEBUG

                    # [TODO]: CHECK IT (BEGIN)
                    # [DONE]: write customized tokenizer: deal with other white space characters (e.g. \t) enclosed by """..."""
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

                    if len(cmd_opt) >= 2:
                        cmd = cmd_opt[0].lower()
                        opt = cmd_opt[1:]
                        logging.debug('   (cmd, opt) = (%s, %s)', cmd, opt) #DEBUG

                        print('Command:  {}'.format(cmd), end='')
                        for item in opt:
                            print('\t"{}"'.format(item), end='')
                        print('')

                        # [DONE]: implement 每一頁可以設定接下來(下一頁開始)的速度
                        if cmd == 'speed':
                            # use opt[0] only
                            # [DONE]: check opt[0] is a float and (float(opt[0]) / 60 >0 and < MAX_FRAME_RATE)
                            if not opt[0].isnumeric():
                                logging.error('   [ERROR] The argument of "speed" command is not a number')
                                pause_exit()
                            if not (0 < float(opt[0]) / 60 < MAX_FRAME_RATE):  # Comparisons can be chained arbitrarily in Python
                                logging.error('   [ERROR] The argument of "speed" command is greater than 0 and less than ' + str(MAX_FRAME_RATE))
                                pause_exit()
                            frameRate[-1] = float(opt[0]) / 60 # frames per minute  -->  frames per second

                        # [TODO]: implement 自動切換語言
                        # [DONE]: implement 自動產生SSML
                        elif cmd == 'say':
                            # use ' '.join(opt)
                            tts_text_file.write(' '.join(opt).strip('\'\"“”‘’`') + '\n')
                            needTTS = True

            # Run TTS
            # [DONE]: use tts_text & tts_audio
            if needTTS:
                # [DONE]: 應該改成根據 workDir (程式安裝路徑)
                tts_exe = os.path.join(ttsDir, 'TTS_engine.exe')
                
                # print('{} "{}" "{}"'.format(tts_exe, cur_tts_text, cur_audio))
                print('Sythesizing narrative...')
                os.system('{} "{}" "{}"'.format(tts_exe, cur_tts_text, cur_audio))  # RELEASE)

                print('InsertAudio,{},{}'.format(pageNum, quotedStr(cur_audio)), file=post_process_script_file)
            else:
                pageAudio[-1] = None



        # [DONE]: (with animation) make page number conform with JPGs
        numPage += 1



# # Step2: Convert PDF to Video
# # Step2.1: Convert PDF to JPGs
# convert_exe = quotedStr(os.path.join(codecDir, 'convert.exe')) #PATH
# os.system(convert_exe + '   -units PixelsPerInch  -density 300 -resize {} {} {}'.format(sizeSpec, pdf, image))


# # Step2.2: Convert JPGs to Video
# # V0.2: with audio in each part of animation
# ffmpeg_exe = quotedStr(os.path.join(codecDir, 'ffmpeg.exe')) #PATH
# with open(mp4list, 'w', encoding = 'UTF-8') as f:
#     for i in range(numPage):
#         cur_audio = audio_specifier.format(i)
#         #cur_audio = audio
#         cur_image = image_specifier.format(i)
#         cur_video = video_specifier.format(i)

#         if numPage == 1: #TODO: hot fix (ad-hoc)
#             cur_image = os.path.join(tmpDir, pdfRoot + '.jpeg')

#         # when audio file exists
#         if pageAudio[i]:
#             os.system(ffmpeg_exe + ' -i {a} -framerate {f} -i {i} -r 30 -y {v}'.format(a=cur_audio, f=frameRate[i], i=cur_image, v=cur_video)) #c:v libx264 -r 30 -pix_fmt yuv420p -y
#         else:
#             os.system(ffmpeg_exe + '        -framerate {f} -i {i} -r 30 -y {v}'.format(f=frameRate[i], i=cur_image, v=cur_video)) #c:v libx264 -r 30 -pix_fmt yuv420p -y
#         f.write('file \'' + cur_video + '\' \n')

# os.system(ffmpeg_exe + ' -f concat -i {l} -c copy -y {v}'.format(l=mp4list, v=video))


pause_exit()