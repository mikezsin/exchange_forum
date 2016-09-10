import win32com.client
import pymysql.cursors
object = win32com.client.Dispatch("Outlook.Application")
ns = object.GetNamespace("MAPI")
tf = ns.GetFolderFromID('<EntryID>') # обращаемся к общей папке с форумом по EntryID
i = 0
tmp = tf.items # получаем все посты/письма
tmp.sort('[ReceivedTime]',False) # сортируем в хронологическом порядке
cnx = pymysql.connect(use_unicode=True, charset='utf8',user='outlook', password='password', host='server',database='outlook')
cursor = cnx.cursor()
for aaa in tmp:
   i +=1
   rtf_file = "post_%d.rtf" %i #задаем имя rtf файла
   if (aaa.Class == 45) or (aaa.Class ==  43): # если объект postitem или mailitem
      aaa.SaveAs('c:\\temp\\low\\store\\%s' %rtf_file ,1) #Save as rtf
      #Извлекаем вложения во временную папку,в моем случае это "c:\temp\low\store\"
      for ac in range(1,aaa.Attachments.Count,1):
         if aaa.Attachments.Item(ac).Type <> 6: # для всех типов, кроме OLE document, с ним пусть разбирается unrtf
            name =  aaa.Attachments.Item(ac).FileName
            ext = name.split('.')[-1]
            filename = 'att_%d_%d.%s' %(i,ac,ext)
            aaa.Attachments.Item(ac).SaveAsFile('c:\\temp\\low\\store\\'+filename)
            cursor.execute("INSERT IGNORE INTO attachments (filename, name, post_id, att_id) VALUES (%s, %s, %s, %s);" ,(filename, name,i,ac))
      #Заносим данные пользователя
      exmail = aaa.SenderEmailAddress
      name = aaa.SenderName
      #Проверяем, есть ли уже такой в таблице users, и если нет, то сразу получаем новый user_id
      cursor.execute("SELECT id FROM users WHERE exmail = '%s' UNION SELECT max(id)+1 FROM users;" %(exmail)) 
      res = cursor.fetchall()
      if len(res) == 2:
         user_id =  res[0][0]
         cursor.execute("UPDATE users SET exist=1 WHERE id=%s;",user_id) #помечаем,что пользователь делал посты 
      elif len(res) == 1:
         #Создаем нового, обязательно задав inab=0, т. к. пользователь неактивный
         user_id =  res[0][0]
         mail = exmail.split('=')[-1]
         if '@' not in mail:
            mail = mail+'@not.exist' #часть писем может быть откуда-то скопирована и к AD отношения не иметь
         cursor.execute("INSERT INTO users (id,exist, exmail,mail, name, inab) VALUES (%s,1, %s, %s, %s, 0);" ,(user_id, exmail,mail,name)) 
      #Разбираемся с темами
      topic = aaa.ConversationTopic # аналогом тем в outlook является ConversationTopic
      tq = """set @mmm = (SELECT IFNULL(max(id), 0)+1 FROM topics);
            INSERT IGNORE INTO
            topics(id,title)
            values (@mmm,%s);"""
      cursor.execute(tq,topic) #заносим в базу
      #Заносим данные в posts
      cursor.execute("SELECT id FROM topics WHERE title = %s;",topic) 
      topic_id =cursor.fetchall()[0][0]
      post_time = int(aaa.ReceivedTime)
      cursor.execute("INSERT IGNORE INTO posts (post_id, user_id, post_time, topic_id, rtf_file) VALUES (%s, %s, %s, %s, %s);",(i,user_id,post_time,topic_id,rtf_file))
   else:
      #Garbage
      cursor.execute("INSERT IGNORE INTO garbage (id, rtf_file,class) VALUES (%s, %s,%s);",(i,rtf_file,aaa.Class))
cursor.close()
cnx.close()
