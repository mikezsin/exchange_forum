#!/usr/bin/python
import lxml.etree, pymysql, subprocess, os
#lxml нам понадобится для работы с тегами

#Пишем функцию для парсинга, cpost - это post_id. /tmp/2del - каталог, куда unrtf выложит извлеченные картинки
def parsertf(cpost):
    p = subprocess.Popen('unrtf /opt/unrtf/store/post_%d.rtf'%cpost, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True,  shell=True, cwd = '/tmp/2del/')
    f = p.stdout.read()
    root = lxml.etree.HTML(f)[1]
    img_id = 0
    for img in root.xpath('//img'): # вложения и картинки определяются как <img>, надо прописать правильные пути
        img_id += 1
        if img.attrib['src'][-4:] == '.wmf': #если img src = *.wmf, - то это вложение.
            #Проверяем, есть ли уже такое извлеченное с помощью outlook
            cursor.execute('SELECT filename,name FROM attachments WHERE post_id=%s AND att_id=%s;'%(cpost,img_id))
            try:
                res = cursor.fetchall()[0]
                img.addnext(lxml.etree.fromstring('<a href="/path/%s">%s</a>'%(res[0],res[1])))
                #Вложение есть, правим ссылку на путь, где оно будет распологатсья на сервере
            except:
                    #Вложения нет, добавляем wmf.
                    attname = 'att_%d_%d'%(cpost,img_id)+img.attrib['src'][-4:]
                    subprocess.Popen('mv -f /tmp/2del/%s /opt/unrtf/store/%s'%(img.attrib['src'], attname), shell=True)
                    img.addnext(lxml.etree.fromstring('<a href="/path/%s">%s</a>'%(attname,attname)))
            img.getparent().remove(img)
        else:
            #Картинки
            imgname = 'img_%d_%d'%(cpost,img_id)+img.attrib['src'][-4:]
            subprocess.Popen('mv -f /tmp/2del/%s /opt/unrtf/store/%s'%(img.attrib['src'], imgname), shell=True)
            img.attrib['src'] = imgname
    root.remove(root[0]) # удаляем шапку письма from/to, и т.п.
    htmltext = lxml.etree.tostring(root)
    cursor.execute('UPDATE posts SET post_text=%s WHERE post_id=%s;',(htmltext,cpost))
    subprocess.Popen('rm -rf /tmp/2del/*', shell=True)
    return htmltext

cnx = pymysql.connect(use_unicode=True, charset='utf8',user='outlook', password='password', host='server',database='outlook')
cursor = cnx.cursor()
#Применяем функцию к каждому посту
cursor.execute('SELECT post_id FROM posts;')
posts =  cursor.fetchall()
for cpost in posts:
    parsertf(cpost[0])

cursor.close()
cnx.close()
