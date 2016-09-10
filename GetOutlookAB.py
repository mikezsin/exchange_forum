import win32com.client # для работы с  Outlook.Application
import pymysql.cursors # для работы с mysql

# создаем COM объект и получаем все AddressEntries
object = win32com.client.Dispatch("Outlook.Application")
ns = object.GetNamespace("MAPI")
als =  ns.AddressLists
gal =  als.Item("Global Address List")
ent =  gal.AddressEntries

#заносим данные в mysql для каждой записи 
cnx = pymysql.connect(use_unicode=True, charset='utf8',user='outlook', password='password', host='server',database='outlook')
cursor = cnx.cursor()
id = 0
for rec in ent:
    id += 1
    exmail = rec.Address # ExchangeUserAddressEntry
    name = rec.Name  # Имя с фамилией
    mail = rec.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E") # получаем аккаунт в виде «account@domain.com»
    cursor.execute("INSERT INTO users (id, mail, exmail, name, inab) VALUES (%s, %s, %s, %s, 1);", (id, mail,exmail,name)) # здесь же заполняем inab =1, т. к. запись присутствует в адресной книге
cursor.close()
cnx.close()
