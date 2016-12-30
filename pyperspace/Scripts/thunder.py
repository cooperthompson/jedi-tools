import os
import sys
import untangle
import requests
import ctypes
import _winreg
import subprocess
import sqlite3


class ThunderEnv:
    def __init__(self):
        self.internal_environments = {}
        self.customer_environments = {}
        appdata = os.getenv('LOCALAPPDATA')
        thunder_path = os.path.join(appdata, 'Epic', 'Thunder')

        thunder_parser = ThunderEnvironmentsParser(thunder_path)
        for env in thunder_parser.environments:
            if env.ini == "DEN":
                self.internal_environments[env.id] = env
            elif env.ini == "ZEN":
                self.customer_environments[env.id] = env

        self.__pre_cache_cust_data()

    def __pre_cache_cust_data(self):
        env_list = (
            (26324, 'chdallas-cmcepicpoc', '4068', 'CHDALLAS_POC', '8.2', 'CMCD POC TEXAS', 'http://hweb'),
            (26323, 'chdallas-cmcepictst', '4072', 'CHDALLAS_TST', '8.2', 'CMCD TST TEXAS', 'http://hweb'),
        )
        connection = sqlite3.connect('pyperspace.db')
        with connection:
            cursor = connection.cursor()
            cursor.execute(
                'CREATE TABLE IF NOT EXISTS Environments(id INT, host TEXT, port INT, epiccomm_id TEXT, version TEXT, display TEXT, hsweb TEXT, CONSTRAINT environment_id_unique PRIMARY KEY(id))')
            for item in env_list:
                cursor.execute('INSERT OR REPLACE INTO Environments VALUES(?,?,?,?,?,?,?)', item)

    def search(self, string):
        results = {}

        # Filter internal environments
        for env_id, env in self.internal_environments.items():
            if string in env.display or string in env.name:
                results[env_id] = env

        # Filter customer environments
        for env_id, env in self.customer_environments.items():
            if string in env.display or string in env.name:
                results[env_id] = env

        # Note that there can be ZEN/DEN id collisions.  For now, I just prioritize matching customer environments
        return results

    def filter(self, string):
        for env_id, env in self.internal_environments.items():
            if string.lower() in env.display.lower() or string.lower() in env.name.lower():
                pass  # match
            else:
                del self.internal_environments[env_id]

        for env_id, env in self.customer_environments.items():
            if string.lower() in env.display.lower() or string.lower() in env.name.lower():
                pass  # match
            else:
                del self.customer_environments[env_id]
        return

    def launch_internal_env(self, den_id):
        env = self.internal_environments.get(den_id)
        if not env:
            ctypes.windll.user32.MessageBoxA(0,
                                             "Environment with ID {} not found.".format(den_id),
                                             "Environment not found", 1)
            return
        env.load_track_data()
        self.launch(env)

    def launch_customer_env(self, zen_id):
        env = self.customer_environments.get(zen_id)
        if not env:
            ctypes.windll.user32.MessageBoxA(0,
                                             "Environment with ID {} not found.".format(zen_id),
                                             "Environment not found", 1)
            return
        env.load_track_data()
        self.launch(env)


    def launch(self, env):
        # cross our fingers and hope the version from DEN matches the registry node
        key = "HKLM:SOFTWARE\Wow6432Node\Epic Systems Corporation\{}".format(env.version)
        reg = _winreg.ConnectRegistry(None, _winreg.HKEY_LOCAL_MACHINE)
        # key = _winreg.OpenKey(key)
        # command = _winreg.OpenKeyEx(key, 'Program')
        # "C:\Program Files (x86)\Epic\v8.3\Shared Files\EpicD83.exe" EDAppServers83.EpicApp Name=Desktop

        ver = env.version.replace(".", "")
        command = r'C:\Program Files (x86)\Epic\v{}\Shared Files\EpicD{}.exe'.format(env.version, ver)
        app_arg = "EDAppServers{}.EpicApp".format(ver)
        env_arg = "Env={}".format(env.epiccomm_id)
        name_arg = "Name={}".format("Desktop{}".format(ver))  # todo: find in registry
        title_arg = "Title={}".format(env.display)

        args = [command, app_arg, env_arg, name_arg, title_arg]
        print " ".join(args)
        subprocess.Popen(args)


class Environment:
    def __init__(self, ini, id, name, customer):
        self.ini = ini  # Should be DEN or ZEN.  DEN = Dev Env  ZEN = Cust Env
        self.id = id
        self.name = name
        self.customer = customer
        self.port = ""
        self.host = ""
        self.version = ""
        self.display = ""
        self.epiccomm_id = ""
        self.hsweb = ""

    def load_track_data(self):
        if self.ini == 'ZEN':
            self.load_customer_data()
            return

        base_url = "http://vs-icx.epic.com/Interconnect-CDE/internal/EDI/HTTP/GetEnvInfo/GetEnvInfo?ID={}"

        try:
            url = base_url.format(self.id)
            response = requests.get(url)
            self.host = response.json().get("EpicCommHost")
            self.port = response.json().get("EpicCommPort")
            self.epiccomm_id = response.json().get("EpicCommID")
            self.version = response.json().get("Version")
            self.display = response.json().get("DisplayName")
            self.hsweb = response.json().get("HSWebURL")
            self.cache_env_data()
        except Exception, e:
            exception_name, exception_value = sys.exc_info()[:2]
            #  Web service failed, check our local cache
            self.load_from_cache()
        finally:
            if self.epiccomm_id == "":
                ctypes.windll.user32.MessageBoxA(0,
                                                 "Web service query failed and environment info was not in cache.\n" +
                                                 "Exception: {}".format(exception_value),
                                                 "Unable to determine EpicComm ID.",
                                                 1)

        if self.epiccomm_id == '':
            ctypes.windll.user32.MessageBoxA(0,
                                             "Unable to get environment info for {}".format(self.id),
                                             "Unable to determine EpicComm ID.",
                                             1)

    def load_customer_data(self):
        '''
        Loads customer data from a pre-cached sqlite database.
        TODO: find better way to get customer server info than manually loading it into the cache.
        :return:
        '''

        connection = sqlite3.connect('pyperspace.db')
        with connection:
            cursor = connection.cursor()
            for row in cursor.execute('SELECT * FROM Environments WHERE id=:id', {'id': self.id}):
                # self.id = row[0]
                self.host = row[1]
                self.port = row[2]
                self.epiccomm_id = row[3]
                self.version = row[4]
                self.display = row[5]
                self.hsweb = row[6]

    def cache_env_data(self):
        env = (
            (self.id, self.host, self.port, self.epiccomm_id, self.version, self.display, self.hsweb)
        )
        connection = sqlite3.connect('pyperspace.db')
        with connection:
            cursor = connection.cursor()
            cursor.execute(
                'CREATE TABLE IF NOT EXISTS Environments(id INT, host TEXT, port INT, epiccomm_id TEXT, version TEXT, display TEXT, hsweb TEXT, CONSTRAINT environment_id_unique UNIQUE (id))')
            cursor.execute('INSERT INTO Environments VALUES(?,?,?,?,?,?)', env)

    def load_from_cache(self):
        connection = sqlite3.connect('pyperspace.db')
        with connection:
            cursor = connection.cursor()
            for row in cursor.execute('SELECT * FROM Environments WHERE id=:id', {'id': self.id}):
                # self.id = row[0]
                self.host = row[1]
                self.port = row[2]
                self.epiccomm_id = row[3]
                self.version = row[4]
                self.display = row[5]
                self.hsweb = row[6]

    def __unicode__(self):
        return '{} - {}[{}] - {}'.format(self.ini, self.name, self.id, self.customer)

    def __str__(self):
        return self.__unicode__()


class ThunderEnvironmentsParser:
    def __init__(self, thunder_path):
        self.thunder_group_file = os.path.join(thunder_path, 'Groups.xml')

        self.environments = []
        self.xmlobj = untangle.parse(self.thunder_group_file)
        self._load()

    def _load(self):
        for list_item in self.xmlobj.ArrayOfEnvironmentListItem.EnvironmentListItem:
            self._parse_env_list_item(list_item)
        pass

    def _parse_env_list_item(self, list_item):
        # Recursively parse.
        # Base case occurs when the Children element has no children
        if not list_item.Children.children:
            rec = list_item.Record
            try:
                customer_name = list_item.CustomerName.cdata if list_item.get_elements("CustomerName") else "Epic"
                env = Environment(list_item.Record.INI.cdata,
                                  list_item.Record.ID.cdata,
                                  list_item.DisplayName.cdata,
                                  customer_name)
            except IndexError as e:
                pass
            self.environments.append(env)
        else:
            self.last_list_item = list_item
            for child_list_item in list_item.Children.EnvironmentListItem:
                # Recurse
                self._parse_env_list_item(child_list_item)


def main(argv):
    thunder_env = ThunderEnv()
    for search_term in argv:
        thunder_env.filter(search_term)

    for env_id in thunder_env.internal_environments.keys():
        thunder_env.launch_internal_env(env_id)
        return

    for env_id in thunder_env.customer_environments.keys():
        thunder_env.launch_customer_env(env_id)
        return

if __name__ == "__main__":
    main(sys.argv[1:])
