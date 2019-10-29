""" Модуль, содержащий библиотеку низкоуровневого взаимодействия с SAP GUI"""

import os, socket, time
import win32gui, win32api, win32con
import win32com.client
import wx

class SAPComRemoteLibrary(object):
    """Библиотека взаимодействия с SAP GUI  """
    sess = None
    sb = None 
    winList = []

    def login(self, syst, uname, pwd, lang="RU"):
        """Инициализирует сессию
        Возвращает 0 в случае успеха, код ошибки в противном случае
        Единственная функция с зашитыми кодами контролов - предполагается, что они не должны
        меняться вообще...
        Метод имеет важный побочный эффект - открытое главное окно приложения
        """

        # проверим - вдруг мы уже залогинены
        if self.sess is not None:
            self.logout()
        
        try:
            sobj = win32com.client.GetObject("SAPGUI")
        except:
            raise AssertionError("Система не доступна (нет SAPGui)".decode("utf-8"))

        try:
                app = sobj.GetScriptingEngine
        except:
            raise AssertionError("Система не доступна (не получается GetScriptingEngine)".decode("utf-8"))

        try:
                conn = app.OpenConnection(syst,True)
        except:
            raise AssertionError("Система не доступна (не получается OpenConnection)".decode("utf-8"))

        if conn.disabledByServer:
            raise AssertionError("На сервере не поддерживается скриптинг".decode("utf-8"))

        try:
            self.sess = conn.sessions[0]
        except:
            raise AssertionError("Не найдены дети у соединения".decode("utf-8"))

        # проверяем, что еще одна сессия работы уже не открыта
        # этого ни разу не видел...
        try:
            nObj = self.sess.findById("wnd[0]/usr/txtRSYST-BNAME")
        except:
            self.sess = None
            raise AssertionError("Некуда вводить логин...".decode("utf-8"))
        nObj.text = uname

        pObj = self.sess.findById("wnd[0]/usr/pwdRSYST-BCODE")
        pObj.text = pwd

        pObj = self.sess.findById("wnd[0]/usr/txtRSYST-LANGU")
        pObj.text = lang

        self.sess.findById("wnd[0]").sendVKey(0)
        self.sb = self.sess.findById("wnd[0]/sbar")

        if len(self.sb.messageType)>0:  # скорее всего - неверный логин или пароль
            sbText = self.get_ctrl_attr("wnd[0]/sbar","text")
            self.sess.findById("wnd[0]").close()
            self.sess = None
            raise AssertionError(sbText)#.decode("windows-1251"))
            # raise AssertionError("Неверный логин или пароль".decode("utf-8"))

        return 0

    def logout(self):
        """Закрывает сессию """

        if self.sess is not None:
            try:
                self.sess.findById("wnd[0]").close()
                self.sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except:
                pass
            self.sess = None
            self.winList = []

        return 0
        
    def send_command ( self, cmd ):
        """ Выполняет команду (транзакцию)
        Возвращает
          - 0 = ок
          - exception с соответствующим текстом в случае ошибки
        """
        if self.sess is None:
            raise AssertionError("Не открыта сессия работы с SAP".decode("utf-8"))

        self.sess.sendCommand( cmd )
        if len(self.sb.messageType)>0:  # такой команды нет (скорее всего)
            raise AssertionError("Команда {0} не существует".decode("utf-8").format(cmd))

        return 0

    def _get_ctrl ( self, uid ):
        """ Получает контрол по его ID, exception в случае ошибки """

        if self.sess is None:
            raise AssertionError("Не открыта сессия работы с SAP".decode("utf-8"))

        try:    # не всегда можно найти по ID...
            ctrl = self.sess.findById(uid)
        except:
            raise AssertionError("Не найден контрол по его ID ({0})".decode("utf-8").format(uid))

        return ctrl

    def show_ctrl ( self, uid ):
        """ Показывает рамку вокруг контрола, exception в случае ошибки """

        if self.sess is None:
            raise AssertionError("Не открыта сессия работы с SAP".decode("utf-8"))

        try:    # не всегда можно найти по ID...
            ctrl = self.sess.findById(uid)
        except:
            raise AssertionError("Не найден контрол по его ID ({0})".decode("utf-8").format(uid))

        ctrl.Visualize(True)

        return 0

    def check_ctrl ( self, uid ):
        """ Проверяет наличие контрола по его ID, возвращает True или False """

        if self.sess is None:
            raise AssertionError("Не открыта сессия работы с SAP".decode("utf-8"))

        try:    # не всегда можно найти по ID...
            ctrl = self.sess.findById(uid)
        except:
            return False

        return True

    def get_ctrl_attr ( self, uid, attr ):
        """ Читает значение контрола (по его имени)
        Возвращает
          - 0+текст = успех
          - exception = ошибка (текст ошибки)
        """
        ctrl = self._get_ctrl(uid)

        try:    # ловим пока все
            retText = getattr( ctrl, attr )#.encode("windows-1251")
        except:
            raise AssertionError("Не получается считать значение атрибута {0} контрола {1}".decode("utf-8").format(attr,uid))

        return retText

    def set_ctrl_attr ( self, uid, attr, astr ):
        """ Заполняет текстовым значением атрибут контрола (по его имени)
        Возвращает
          - 0 = успех
          - код ошибки
        """
        
        ctrl = self._get_ctrl(uid)

        try:    # пока не понимаем, какие могут быть эксепшены...
            setattr( ctrl, attr, astr )
        except:
            raise AssertionError("Не получается заполнить атрибут {2} контрола {0} текстом ({1})".decode("utf-8").format(uid,astr,attr))

        return 0

    def set_ctrl_spaced_attr ( self, uid, attr, astr, alen ):
        """ Заполняет текстовым значением атрибут контрола (по его имени), дополняет слева пробелами до длины alen
        Возвращает
          - 0 = успех
          - код ошибки
        """

        ctrl = self._get_ctrl(uid)
        aVal = " "*(int(alen)-len(astr)) + astr

        return self.set_ctrl_attr(uid, attr, aVal)
    
    def exec_ctrl_func ( self, uid, func, *params ):
        """ Выполняет метод контрола по его имени
        Возвращает
          - 0 = успех
          - exception = ошибка (код ошибки)
        """
        ctrl = self._get_ctrl(uid)

        try:    # не все контролы умеют нажиматься
            getattr( ctrl, func )( *params )
        except:
            raise AssertionError("Не получилось выполнить функцию {0} для контрола {1}".decode("utf-8").format(func,uid))

        return 0

    def select_row (self, uid, rowNo ):
        """ выделяет строку в табличном контроле (TableControl) """

        ctrl = self._get_ctrl(uid)

        try:    # не все контролы умеют нажиматься
            aRow = ctrl.getAbsoluteRow(rowNo)
            aRow.selected = True
        except:
            raise AssertionError("Не получилось выделить строку {0} для контрола {1}".decode("utf-8").format(rowNo,uid))

    def get_cell_value ( self, uid, rowno, colname ):
        """ Читает значение ячейки в табличном контроле
        Возвращает
          - текст = успех
          - exception = ошибка (текст ошибки)
        """
        ctrl = self._get_ctrl(uid)

        try:    # ловим пока все
            retText = ctrl.getCellValue( rowno, colname )#.encode("windows-1251")
        except:
            raise AssertionError("Не получилось считать значение колонки {0} контрола {1} в строке {2}".decode("utf-8").format(colname,uid,rowno))

        return retText
    
    def exec_menu_command ( self, itemno ):
        """ Выполняет команду системного меню (которое по ALT+SPACE), 
            параметр - сколько раз нужно нажать стрелку вниз, для того, чтобы нажатие на ENTER выполнило именно эту команду """
        
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.AppActivate(self.get_ctrl_attr("wnd[0]","text"))      # таким способом выбирается окно, в котором нужно выполнить команду
        shell.SendKeys("% ")                                        # ALT+SPACE
        for i in range(0,int(itemno)):
                shell.SendKeys("{DOWN}")
        shell.SendKeys("{ENTER}")

    def screenForCtrl ( self, uid, fname ):
        """ скриншот контрола по его ID """

        ctrl = self._get_ctrl(uid)          # находим контрол
        return self.makeScrShot ( ctrl.ScreenLeft, ctrl.ScreenTop, ctrl.Width, ctrl.Height, fname )

    def makeScrShot( self,x,y,w,h,fn,useIrfan=True ):
        """ техническое делание скриншота (по координатам) 
            координаты используются только не irfanview, это только для скриншотов для документации
        """

        wasFile = True
        if useIrfan:
            i_view = [ "c:\Program Files\IrfanView\i_view64.exe",
                       "c:\Program Files\IrfanView\i_view32.exe",
                       "c:\Program Files (x86)\IrfanView\i_view32.exe",
                       "c:\Program Files (x86)\IrfanView\i_view64.exe" ]
            for iv in i_view:
                if os.path.exists(iv):
                    subprocess.call([iv,"/capture=0","/convert=%s"%(fn)])
                    break
            else:
                wasFile = False
                
        else:
            x = int(x)
            y = int(y)
            w = int(w)
            h = int(h)
            app = wx.App(False)
            dcScreen = wx.ScreenDC()
            #Create a Bitmap that will hold the screenshot image later on
            #Note that the Bitmap must have a size big enough to hold the screenshot
            bmp = wx.EmptyBitmap(w, h)
     
            #Create a memory DC that will be used for actually taking the screenshot
            memDC = wx.MemoryDC()
     
            #Tell the memory DC to use our Bitmap
            #all drawing action on the memory DC will go to the Bitmap now
            memDC.SelectObject(bmp)
     
            #Blit (in this case copy) the actual screen on the memory DC
            #and thus the Bitmap
            memDC.Blit( 
                0, #Copy to this X coordinate
                0, #Copy to this Y coordinate
                w, #Copy this width
                h, #Copy this height
                dcScreen, #From where do we copy?
                x, #What's the X offset in the original DC?
                y  #What's the Y offset in the original DC?
            )
     
            #Select the Bitmap out of the memory DC by selecting a new
            #uninitialized Bitmap
            memDC.SelectObject(wx.NullBitmap)
            img = bmp.ConvertToImage()
            img.SaveFile(fn, wx.BITMAP_TYPE_PNG)

        if wasFile:	# если у нас получилось сформировать файл
            # read file data 
            fp = open(fn,"rb")
            buf = fp.read()
            fp.close()
            # возвращаем бинарные данные - работает (изначально пробовал конвертить в 16ичную строку)
            return buf
        else:
            return ""

    def pause_execution ( self, txt="Execution paused...", mode=0 ):
        """ преостанавливает выполнение теста, ждет нажатия на ОК (для целей отладки тестов) """

        app = wx.App(False)
        if mode==0:     # просто кнопка ОК
            dlg = wx.MessageDialog(None, txt, 'Paused', wx.OK|wx.ICON_INFORMATION)
            dlg.ShowModal()
            res = 0
        else:            # OK=1 Cancel=0 и возвращаем эти числа
            # dlg = wx.MessageDialog(None, txt, 'Ответьте на вопрос'.decode("utf-8"), wx.OK|wx.CANCEL|wx.ICON_QUESTION|wx.STAY_ON_TOP) - последние 2 не работают...
            dlg = wx.MessageDialog(None, txt, 'Ответьте на вопрос'.decode("utf-8"), wx.YES|wx.NO)
            res = 5104 - dlg.ShowModal()        # магическая константа...
        dlg.Destroy()
        return res

    def debug_execution ( self ):
        """ преостанавливает выполнение теста, ждет ввода имени переменной или Cancel для завершения режима отладки """

        app = wx.App(False)
        dlg = wx.TextEntryDialog(None, 'Введите имя переменной'.decode("utf-8"), 'Debug Window')
        dlg.ShowModal()
        res = dlg.GetValue()	# все равно, нажмет ли пользователь Cancel или просто ничего не введет
        if res:
            res = "\${" + res + "}"
        dlg.Destroy()
        return res

    def save_sut_file ( self, fname, fdata ):
        """ получает данные, сохраняет их в файле по нужному пути """

        fp = open(fname,"wb")
        buf = fp.write(fdata)
        fp.close()

        return "OK"

def getMyIP(remoteip="8.8.8.8"):
    """ get my IP address in a strange way... """

    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(("8.8.8.8", 80))
    return s.getsockname()[0]

if __name__ == '__main__':

    from robotremoteserver import RobotRemoteServer

    # порт - либо идет после точки в имени текущей директории, либо стандартный (8270)
    port = os.path.splitext(os.getcwd())[1]	# будет либо пусто, либо порт с точкой вначале
    if port:
        RobotRemoteServer(SAPComRemoteLibrary(), host=getMyIP(),port=port.split(".")[1])
    else:
        RobotRemoteServer(SAPComRemoteLibrary(), host=getMyIP())


