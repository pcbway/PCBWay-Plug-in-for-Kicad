#https://opensource.org/licenses/MIT 

import os
import webbrowser
import shutil
import requests
import wx
import tempfile
from threading import Thread
from .result_event import *
from .config import *
from .process import *


class PCBWayThread(Thread):
    def __init__(self, wxObject):
        Thread.__init__(self)
        self.process = PCBWayProcess()
        self.wxObject = wxObject
        self.start()

    def run(self):
        
        temp_dir = tempfile.mkdtemp()
        _, temp_file = tempfile.mkstemp()

        try:
            self.report(5)

            self.process.get_gerber_file(temp_dir)

            self.report(15)

            self.process.get_netlist_file(temp_dir)

            self.report(25)

            self.process.get_components_file(temp_dir)

            self.report(35)

            gerberData = self.process.get_gerber_parameter()

            self.report(45)

            p_name = self.process.get_name()

            temp_file = shutil.make_archive(p_name, 'zip', temp_dir)
            files = {'upload[file]': open(temp_file, 'rb')}

            self.report(55)

            upload_url = baseUrl + '/Common/KiCadUpFile/'
            
            self.report(65)
            
            rsp = requests.post(
                upload_url, files=files, data={'boardWidth':gerberData['boardWidth'],'boardHeight':gerberData['boardHeight'],'boardLayer':gerberData['boardLayer']})
            
            self.report(75)

            urls = json.loads(rsp.content)

            readsofar = 0
            totalsize = os.path.getsize(temp_file)
            with open(temp_file, 'rb') as file:
                while True:
                    data = file.read(10)
                    if not data:
                        break
                    readsofar += len(data)
                    percent = readsofar * 1e2 / totalsize
                    self.report(75 + percent / 9)

        except Exception as e:
            wx.MessageBox(str(e), "Error", wx.OK | wx.ICON_ERROR)
            self.report(-1)
            return
       

        webbrowser.open(urls['redirect'])
        self.report(-1)

    def report(self, status):
        wx.PostEvent(self.wxObject, ResultEvent(status))
        



    