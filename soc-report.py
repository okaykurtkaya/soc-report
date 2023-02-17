# ########################################################
import os                                                #
import sys                                               #
import re                                                #
import warnings                                          #
import pandas as pd                                      #
from docxtpl import DocxTemplate                         #
import docx                                              #
from datetime import datetime                            #
import redminelib                                        #
import time                                              #
import locale                                            #
from docx.shared import Mm                               #
from docxtpl import InlineImage                          #
import matplotlib.pyplot as plt                          #
import glob                                              #
import requests                                          #
# ########################################################
class MonthlyReport:
    def __init__(self):
        self.main_url = "https://ticket.<YourTicketSystem>.com/"

        # FOR L2 INVENTORY LIST
        self.zabbix_customers = {
            "Customer - 1" : ["https://FirstCustomer.zabbix.local/api_jsonrpc.php", "<FirstCustomer_zabbix_token>"],
            "Customer - 2" : ["http://SecondCustomer.zabbix.local/zabbix/api_jsonrpc.php", "<SecondCustomer_zabbix_token>"],
            "Customer - 3" : ["<zabbix_url>", "<zabbix_token>"],
            # IF YOU HAVE MORE..
        }
        
        # DETAILS OF INVENTORY LIST
        self.hosts_info = []
        
        # LOGIN TICKET SYSTEM
        self.username = input("ticket username > ")
        self.password = input("ticket password > ")
        
        # REQUEST TO TICKET SYSTEM
        if not sys.warnoptions:
            warnings.simplefilter("ignore") # FOR SSL WARNING. I IGNORED..
            self.redmine = redminelib.Redmine(self.main_url, requests = {'verify' : False}, username = self.username, password = self.password)
            """
            IF YOU DON'T KNOW HOW CAN I REQUEST TO TICKET SYSTEM, 
            YOU SHOULD TO -> https://python-redmine.com/configuration.html
            """
        # IS USER LOG IN ?
        if not self.redmine.auth():
            print(f"\n[-] INVALID USERNAME OR PASSWORD") 
            time.sleep(1.3)
            sys.exit()
        
        # FOR EACH CUSTOMER
        self.user_project = ""

        # CUSTOMERS IN TICKET SYSTEM
        self.projects = list(self.redmine.project.all())

        # PROCESS OF DATE
        self.start_date_input = input("start date (Year-Month-Day | exs; 2022-1-1) > ").split("-")
        self.start_date = datetime.strftime(datetime(int(self.start_date_input[0]), int(self.start_date_input[1]), int(self.start_date_input[2])), "%Y-%m-%d")
        
        self.last_date_input = input("last date (Year-Month-Day | exs; 2022-1-31) > ").split("-")
        self.last_date = datetime.strftime(datetime(int(self.last_date_input[0]), int(self.last_date_input[1]), int(self.last_date_input[2])), "%Y-%m-%d")

        # BASE DOCX TEMPLATE
        self.base_report_file = input("base template docx file > ")

    # BASE REQUEST FOR ALL CUSTOMERS
    def project_base_settings(self, project_name):
        # GETS ALL CUSTOMER INTO self.project
        for self.project in self.projects:
            self.user_project = project_name
            if self.project.name == self.user_project:
                self.issues = self.redmine.issue.filter(
                    project_id = self.project.id,
                    created_on = f'><{self.start_date}|{self.last_date}',
                    status_id = '*',
                    sort = 'category:desc'
                )
                """
                IF YOU DON'T KNOW HOW CAN I ADD FILTER TO MY REQUEST,
                YOU SHOULD TO -> https://python-redmine.com/resources/issue.html
                """
                # TOTAL NUMBER OF TICKET
                self.total_ticket_number = len(self.issues)
    
    # CATEGORY OF EACH TICKET IN TICKET SYSTEM
    def trackers(self, trackerID, project_name):
        for self.project in self.projects:
            self.user_project = project_name
            if self.project.name == self.user_project:
                self.issues = self.redmine.issue.filter(
                    project_id = self.project.id,
                    created_on = f'><{self.start_date}|{self.last_date}',
                    tracker_id = int(trackerID),
                    status_id = '*',
                    sort = 'category:desc'
                )

                # IN THE ABOVE QUERY, I BRING THE TICKETS THAT WERE OPENED FOR ALL CATEGORIES.

                # WILL BE USED FOR EACH CUSTOMER
                self.play_text = []
                # FOR CHARTS
                self.context = [set()]
    # CREATES EXCEL FILE FOR VALUE'S CONTROL OF SOME DANGEROUS_RATE_STATES, WINDOWS_SECURITY_EVENT, ATTACK_VECTORS CATEGORIES
    def control_trackers_excel(self, dataframe, func_name, project_name, sort_parameter):
        if len(dataframe) > 0: # IF DATA IN DATAFRAME..
            print(f"\t[+] EVENTS CREATED TO > {project_name}/{func_name} <\n") 
            # IF YOU'VE ANY OPENED WRONG TICKET IT CREATES EXCEL FILE FOR CONTROL.
            dataframe.sort_values(by = str(sort_parameter), ascending = False).to_excel(f"{project_name}/{str(project_name).lower().replace(' ', '_')}_{func_name}.xlsx")
        else:
            print(f"\t[-] I COULDN'T FIND ITEM IN {project_name}-{func_name.upper().replace('_', ' ')}\n")
    """
        THIS FUNCTION IS CREATED FOR CONTROL PURPOSES. SOME TIMES, 
        IN CASE THE WRONG TICKET IS OPENED IN THE TICKET SYSTEM, 
        EXCEL FILE IS CREATED FOR CONTROL PURPOSES. IF THERE IS SOMETHING WRONG, 
        YOU MUST ADD IT STATICLY, CREATING EXCEL CHART.
    """
    # TO INFORM FOR THE PERSON RUNNING THE PROGRAM
    def control_trackers_word(self, dataframe, func_name, project_name):
        if len(dataframe) > 0: # IF DATA IN DATAFRAME..
            print(f"\t[+] {project_name}/{func_name.lower().replace('_', ' ')} ARE WRITTEN IN THE REPORT FILE..\n")
        else:
            print(f"\t[-] I COULDN'T FIND ITEM IN > {project_name}/{func_name.lower().replace('_', ' ')} <\n") 

    # SUB CATEGORY IN CATEGORY_ID = 6 
    def system_performance_statistics_source_ips(self):
        # TRACKER_ID IS A CATEGORY IN TICKET SYSTEM.
        self.trackers(trackerID = 6, project_name = self.user_project)

        source_ips_text = ""

        for issue in self.issues:
            # SUBMITS ALL IP ADDRESSES OF THIS CATEGORY TO THE VARIABLE source_ips_text.
            source_ips_text += "\n" + str(issue.custom_fields[0].value)
        
        # CREATES A DATEFRAME FOR IP's
        self.system_performance_source_ip = pd.DataFrame(re.findall(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', source_ips_text))

        self.system_performance_source_ip_control_number = len(self.system_performance_source_ip)

        # WRITES ALL IP's TO SOC DOCX FILE..
        self.control_trackers_word(self.system_performance_source_ip, self.system_performance_statistics_source_ips.__name__, self.user_project)
    
    # SUB CATEGORY IN CATEGORY_ID = 6
    def system_performance_statistics(self): # SIMILAR PROCESSES AS ABOVE..
        self.trackers(trackerID = 6, project_name = self.user_project)

        self.system_events_ticket_number = len(self.issues)

        for issue in self.issues:
            self.play_text.append(str(issue.custom_fields[1].value))
        
        self.system_performance = pd.DataFrame(self.play_text)

        self.system_performance_control_number = len(self.system_performance)

        self.control_trackers_word(self.system_performance, self.system_performance_statistics.__name__, self.user_project)
    
    # SUB CATEGORY IN CATEGORY_ID = 2
    def dangerous_rate_state(self):
        self.trackers(trackerID = 2, project_name = self.user_project)

        for issue in self.issues:
            # CUSTOM FIELD INSIDE EACH TICKET. HERE IS ONLY FOR HARMFUL IP CONDITION..
            if str(issue.custom_fields[7].value) != "": # NONE VALUE. I DON'T WANT THIS..
                self.play_text.append(str(issue.custom_fields[7].value))
        
        for item in self.play_text:
            if item == "1": # HARMFUL IP
                self.context[0].add(("RATE OF HARMFUL IP", self.play_text.count(item)))
            elif item == "0": # NO HARMFUL IP
                self.context[0].add(("RATE OF NO HARMFUL IP", self.play_text.count(item)))

        df = pd.DataFrame(self.context[0], columns = ["LABELS", "VALUE OF LABELS"])
        
        self.dangerous_rate_state_control_number = len(self.context[0])
        
        if len(self.context[0]) > 0:
            self.control_trackers_excel(df, self.dangerous_rate_state.__name__, self.user_project, "VALUE OF LABELS")

            # CREATES A PIE CHART FOR RATE OF IP
            self.draw_pie_chart(df["LABELS"].values, df["VALUE OF LABELS"].values, self.user_project, self.dangerous_rate_state.__name__)
    
    # SUB CATEGORY IN CATEGORY_ID = 2
    def target_ports(self):
        self.trackers(trackerID = 2, project_name = self.user_project)

        target_ports_text = ""

        for issue in self.issues:
            # ALL ISSUES IN THE TICKET SYSTEM ARE RECEIVED IN A CYCLE AND ADDED TO THE RELATED PART IN WHICH FIELD IN THE JSON FILE
            # THIS VALUE MAY VARY ON YOUR TICKET SYSTEM SO THAN YOU CAN CHECK JSON INFO OF TICKET.
            target_ports_text += "\n" + str(issue.custom_fields[5].value)
        
        # REGEX RULE
        self.destination_port = pd.DataFrame(re.findall(r'(\b\d{2,5}\b)', target_ports_text))

        # FOR DOCX FILE
        self.target_ports_control_number = len(self.destination_port)

        # FOR USER LOG
        self.control_trackers_word(self.destination_port, self.target_ports.__name__, self.user_project)
    
    # SUB CATEGORY IN CATEGORY_ID = 2
    def locations(self):
        self.trackers(trackerID = 2, project_name = self.user_project)

        for issue in self.issues:
            self.play_text.append(str(issue.custom_fields[3].value))
        
        self.location = pd.DataFrame(self.play_text)

        self.location_control_number = len(self.location)

        self.control_trackers_word(self.location, self.locations.__name__, self.user_project) 
    
    # SUB CATEGORY IN CATEGORY_ID = 2
    def target_ips(self):
        self.trackers(trackerID = 2, project_name = self.user_project)

        destination_text = ""

        for issue in self.issues:
            destination_text += "\n" + str(issue.custom_fields[2].value)
        
        self.target_ip = pd.DataFrame(re.findall(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', destination_text))

        self.target_ips_control_number = len(self.target_ip)

        self.control_trackers_word(self.target_ip, self.target_ips.__name__, self.user_project) 

    # SUB CATEGORY IN CATEGORY_ID = 2
    def source_ips(self):
        self.trackers(trackerID = 2, project_name = self.user_project)
        
        source_ips_text = ""

        for issue in self.issues:
            source_ips_text += "\n" + str(issue.custom_fields[1].value)
        
        self.source_ip = pd.DataFrame(re.findall(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', source_ips_text))

        self.source_ips_control_number = len(self.source_ip)
    
        self.control_trackers_word(self.source_ip, self.source_ips.__name__, self.user_project)
    
    # SUB CATEGORY IN CATEGORY_ID = 8
    def windows_security_events(self):
        self.trackers(trackerID = 8,  project_name = self.user_project)

        self.windows_events_ticket_number = len(self.issues)

        self.windows_events_id = ""

        for issue in self.issues:
            self.windows_events_id += "\n" + str(issue.custom_fields[4].value)

        # IN MY TICKET SYSTEM, t COMES NEXT TO WINDOWS EVENT IDs. THAT'S WHY I APPLIED THE REPLACE PROCESS..
        # IF YOU HAVE DIFFERENT EVENT'S ID YOU CAN CHANGE REGEX RULE..
        events_id = re.findall(r'\b\d{4}\b|\bOff\w*\b|\bA\w*\b|\bDiğer\b', self.windows_events_id.replace("t", " "))

        if events_id:
            for event in events_id:
                self.context[0].add((event, events_id.count(event)))
        
        self.windows_events = pd.DataFrame(self.context[0], columns = ["WINDOWS SECURITY EVENTS", "NUMBER OF PROBLEMS EXPERIENCED"])

        if len(self.context[0]) > 0:
            self.control_trackers_excel(self.windows_events, self.windows_security_events.__name__, self.user_project, "NUMBER OF PROBLEMS EXPERIENCED")

            self.draw_pie_chart(self.windows_events["WINDOWS SECURITY EVENTS"].values, self.windows_events["NUMBER OF PROBLEMS EXPERIENCED"].values, self.user_project, self.windows_security_events.__name__)

    # SUB CATEGORY IN CATEGORY_ID = 2
    def attack_vectors(self):
        self.trackers(trackerID = 2, project_name = self.user_project)

        self.incident_events_ticket_number = len(self.issues)

        for issue in self.issues:
            self.play_text.append(str(issue.custom_fields[0].value))

        for item in self.play_text:
            self.context[0].add((item, self.play_text.count(item)))
                
        df = pd.DataFrame(self.context[0], columns = ["ATTACK VECTORS", "NUMBER OF PROBLEMS EXPERIENCED"])

        if not os.path.exists(f"{self.user_project}"):
            os.mkdir(f"{self.user_project}")

        if len(self.context[0]) > 0:
            self.control_trackers_excel(df, self.attack_vectors.__name__, self.user_project, "NUMBER OF PROBLEMS EXPERIENCED")

            if 2.2 < len(df): # FOR WIDTH OF Y-VALUES ON CHART
                self.horizontal_chart(df, "ATTACK VECTORS", "NUMBER OF PROBLEMS EXPERIENCED", 0.03)
            else:
                self.horizontal_chart(df, "ATTACK VECTORS", "NUMBER OF PROBLEMS EXPERIENCED", 0.9)
    
    """
        WHEN WE EXAMINED THE ABOVE PROCESSES, MANY OF THESE ARE CREATED IN A SIMILAR WAY. 
        I DON'T WANT TO EXPLAIN MUCH ABOUT THAT. I HOPE YOU UNDERSTAND..
    """
    
    # DRAW A HORIZONTAL CHART FOR ATTACK VECTORS
    def horizontal_chart(self, dataframe, y_val, w_val, distance_of_values_in_y_line):
        fig, ax = plt.subplots(figsize=(9, 5.5))

        y = dataframe[y_val].values
        w = dataframe[w_val].values

        ax.barh(y, w, 0.35)
        ax.set_yticks(y, labels = y, fontsize = 9)
        ax.set_frame_on(False)
        fig.subplots_adjust(left = 0.39, right = 0.99)
        ax.set_ymargin(distance_of_values_in_y_line)
        ax.set_axisbelow(True)
        ax.xaxis.grid(visible = True)

        plt.savefig(f"{self.user_project}/{str(self.user_project).lower().replace(' ', '_')}_{self.attack_vectors.__name__}.png")

    # DRAW A CHART FOR DANGEROUS RATE STATES AND WINDOWS SECURITY EVENTS 
    def draw_pie_chart(self, y_val, w_val, project_name, func_name):
        colors = [
        "#F44336", "#9C27B0", "#673AB7", "#3F51B5", "#2196F3", "#E91E63", 
        "#00BCD4", "#009688", "#FFEB3B", "#FFC107", "#4CAF50", "#D1F2EB",
        "#8BC34A", "#795548", "#9E9E9E", "#607D8B", "#FFCCBC", "#B3E5FC", 
        "#222222", "#1A237E", "#827717", "#880E4F", "#00FFFF", "#FF00FF"]

        fig1, ax1 = plt.subplots()

        ax1.pie(w_val, labels = y_val, colors = colors, textprops = {'fontsize': 6.5}, labeldistance = 65, wedgeprops = {'linewidth' : 0.9, 'edgecolor' : 'white'}) # startangle = 90, autopct='%1.1f%%', pctdistance=1.11 FOR %

        box = ax1.get_position()

        ax1.set_position([box.x0, box.y0 + box.height * 0.1, box.width, box.height * 1])
            
        ax1.legend(loc = 'upper center', bbox_to_anchor = (0.50, 0.01), fancybox = True, ncol = 6, fontsize = 7.5)

        plt.savefig(f"{project_name}/{str(project_name).lower().replace(' ', '_')}_{func_name}.png")

        """
            IF YOU WANT TO KNOW HOW TO CREATE A GRAPH, YOU SHOULD LOOK HERE 
            
            > https://matplotlib.org/stable/gallery/pie_and_polar_charts/pie_features.html <

            > https://proclusacademy.com/blog/customize_matplotlib_piechart/ <
        """

    # WRITE DIFFERENTLY CONTENTS OF CUSTOMER TO DOCX FILE
    def write_docx(self, project_name):
        """
            IF YOU WANT TO SEE HOW PYTHON AND DOCX FILES ARE USED, 
            LOOK HERE > https://docxtpl.readthedocs.io/en/latest/ <
        """
        locale.setlocale(locale.LC_ALL, "")

        doc = DocxTemplate(self.base_report_file)

        # IF FILE TYPE IS .PNG GET IT
        customer_logo = glob.glob('images/' + f"{str(project_name).lower().replace(' ', '_').replace('ı', 'i').replace('İ', 'i').replace('ü', 'u').replace('Ü', 'u').replace('ğ', 'g').replace('Ğ', 'g').replace('ö', 'o').replace('Ö', 'o').replace('ş', 's').replace('Ş', 's')}*")
        
        # THE VALUES TO BE SENT TO THE CONTEXT REPORT FILE..
        context = {
            "month": datetime.strftime(datetime(int(self.last_date_input[0]), int(self.last_date_input[1]), int(self.last_date_input[2])), '%B'),
            "year": self.last_date_input[0],
            "user_project": project_name,
            "logo": InlineImage(doc, customer_logo[0], width=Mm(80), height=Mm(30)),
            "attack_vectors_graph": InlineImage(doc, f"{project_name}/{str(project_name).lower().replace(' ', '_')}_{self.attack_vectors.__name__}.png", width=Mm(160)),
            "windows_security_events_graph": InlineImage(doc, f"{project_name}/{str(project_name).lower().replace(' ', '_')}_{self.windows_security_events.__name__}.png", width=Mm(160)),
            "dangerous_state_graph": InlineImage(doc, f"{project_name}/{str(project_name).lower().replace(' ', '_')}_{self.dangerous_rate_state.__name__}.png", width=Mm(160)),
            "total_ticket_number": self.total_ticket_number,
            "incident_info_ticket_number": self.incident_events_ticket_number,
            "system_performance_ticket_number": self.system_events_ticket_number,
            "windows_events_ticket_number": self.windows_events_ticket_number,
            "dangerous_state_control_number": self.dangerous_rate_state_control_number,
            "target_ports_control_number": self.target_ports_control_number,
            "source_ips_control_number": self.source_ips_control_number,
            "target_ips_control_number": self.target_ips_control_number,
            "location_control_number": self.location_control_number,
            "system_performance_control_number": self.system_performance_control_number,
            "system_performance_source_ip_control_number": self.system_performance_source_ip_control_number,
            "source_ips": list(self.source_ip.value_counts().items()),
            "target_ips": list(self.target_ip.value_counts().items()),
            "countries": list(self.location.value_counts().items()),
            "windows_events": list(self.windows_events.sort_values(by="Yaşanan Sorun Sayısı", ascending=False).values),
            "destination_ports": list(self.destination_port.value_counts().items()),
            "system_performance": list(self.system_performance.value_counts().items()),
            "system_performance_source_ip": list(self.system_performance_source_ip.value_counts().items()),
            "hosts_info": self.hosts_info,
        }
        doc.render(context) # # SUBMITS ALL VALUES
        # REPORTS ARE CREATED IN A SEPARATE FOLDER FOR EACH CUSTOMER.
        doc.save(f"{project_name}/{project_name} - SOC Service Report - {datetime.strftime(datetime(int(self.last_date_input[0]), int(self.last_date_input[1]), int(self.last_date_input[2])), '%B')} {self.last_date_input[0]}.docx")        
# #############################################################################################################
def entry_to_system():
    title = """
    __  _______  _   __________  ______  __            ____  __________  ____  ____  ______
   /  |/  / __ \/ | / /_  __/ / / / /\ \/ /           / __ \/ ____/ __ \/ __ \/ __ \/_  __/
  / /|_/ / / / /  |/ / / / / /_/ / /  \  /  ______   / /_/ / __/ / /_/ / / / / /_/ / / /   
 / /  / / /_/ / /|  / / / / __  / /___/ /  /_____/  / _, _/ /___/ ____/ /_/ / _, _/ / /    
/_/  /_/\____/_/ |_/ /_/ /_/ /_/_____/_/           /_/ |_/_____/_/    \____/_/ |_| /_/ 
\nAuthor: n3gat1v3o\nVersion: 1.1\nNotForYou: Search For > BUG < On Me :)"""

    for i in range(3):
        os.system("color d")
        print(f"{title}\n{'_' * 85}\n")
        time.sleep(0.3)
        os.system("color f")
        time.sleep(0.3)
        os.system("cls")
    
    print(f"{title}\n{'_':_>97}\n")
    """
        PROGRAM INTRODUCTION. BECAUSE IT IS MY FIRST PROGRAM, I MAY HAVE DECORATED A LITTLE TOO MUCH ¿
    """
# #############################################################################################################
def main():
    try:
        entry_to_system()
        result = MonthlyReport()
        # THESE PROCESSES ARE MADE FOR ALL CUSTOMERS IN THE TICKET SYSTEM..
        for item in result.projects:
            if str(item) != "Customer Closed" and str(item) != "EXAMPLE - D CUSTOMER":
                """
                IF YOU DON'T WANT TO MAKE A REPORT FOR A CUSTOMER, 
                YOU CAN WRITE THE NAME ON THE TICKET SYSTEM OR REMOVE IT FROM THE TICKET SYSTEM. 
                SOME CUSTOMERS SHOULD STAY IN THE SYSTEM I USE. THAT'S WHY I ADDED SUCH CONTROL.
                """
                os.system("cls")
                # THE SECTION WHERE THE REPORTS ARE CREATED..
                print(f"> {str(item).upper()} REPORTS ARE CREATING..\n")
                result.project_base_settings(str(item))
                result.attack_vectors()
                result.windows_security_events()
                result.system_performance_statistics()
                result.source_ips()
                result.target_ips()
                result.locations()
                result.target_ports()
                result.dangerous_rate_state()
                result.system_performance_statistics_source_ips()
                
                # INVENTORY LIST FOR ZABBIX
                if str(item) in result.zabbix_customers:
                    get_hosts = requests.post(result.zabbix_customers[str(item)][0], json = {
                    "jsonrpc": "2.0",
                    "method": "host.get",
                    "params": {
                        "output": [
                            "name",
                            "active_available",
                            "available"
                        ],
                        "selectInterfaces": [
                            "ip",
                            "available"
                        ]
                    },
                    "id": 1,
                    "auth": result.zabbix_customers[str(item)][-1]
                }, verify = False)
                    """
                        IF YOU WANT TO MAKE A REQUEST BY ZABBIX, 
                        YOU CAN LOOK HERE -> https://www.zabbix.com/documentation/current/en/manual/api/reference/host/get
                    """

                    for vals in get_hosts.json()['result']:
                        if str(item) == "Customer1" or str(item) == "Customer2":
                            # IF ZABBIX SERVER VERSION IS SMALL THAN 5.4
                            for iface in vals['interfaces']:
                                if iface['ip'] != "127.0.0.1":
                                    result.hosts_info.append({"host" : str(vals['name']), "ip" : iface['ip'], "available" : str(iface['available'])})
                        else:
                            # IF ZABBIX SERVER VERSION IS LARGER THAN 5.4
                            for iface in vals['interfaces']:
                                if iface['ip'] != "127.0.0.1":
                                    result.hosts_info.append({"host" : str(vals['name']), "ip" : iface['ip'], "available" : str(vals['available'])})
                
                    """
                        VALUES IN CYCLES MAY VARY ON YOUR ZABBIX SYSTEMS. 
                        IF YOU DON'T WANT THIS, YOU CAN REMOVE THIS SECTION 
                        BUT REMEMBER IF YOU ARE REMOVING SOMETHING, YOU MUST REMOVE IT 
                        FROM BOTH THE CODE AND THE DOCX FILE.
                    """
                # SORT ALL SERVERS..
                result.hosts_info.sort(key=lambda item: item['available'])
                # ALL PROCESSES ARE OK.
                print(f"[+] {str(item).upper()} FILES ARE PREPARED.\n")
                # WRITES TO SOC REPORT DOCX FILE..
                result.write_docx(str(item))
                # THE LIST MUST BE CLEARED FOR EACH CUSTOMER..
                result.hosts_info.clear()
    # ERROR HANDLINGS..
    except PermissionError:
        print("\n[-] WHEN PROGRAM IS RUNNING IF YOUR ANY EXCEL|WORD FILE IS OPEN, PROGRAM WASN'T UPDATE.")
        sys.exit()
    except requests.exceptions.ConnectTimeout:
        print(f"\n[-] PROGRAM DID NOT WORK BECAUSE YOU SHOULD CONNECT TO VPN !")
        sys.exit()
    except requests.exceptions.ConnectionError:
        print(f"\n[-] CONNECTION FAILD.")
        sys.exit()
    except OSError:
        print("\n[-] CAN NOT SAVE FILE INTI A NON-EXISTENT DIRECTORY !")
        sys.exit()
    except KeyboardInterrupt:
        print("\n[-] KEY DETECTED SO THAN PROGRAM IS STOPPED.")
        sys.exit()
    except AttributeError as err:
        print(f"\n[-] ATTRIBUTE ERROR > {str(err).upper()} <")
        input()
    except docx.opc.exceptions.PackageNotFoundError:
        print("\n[-] BASE PATERN DOCX FILE NOT FOUND, PLEASE ENTER CORRECTLY FILE NAME OR MOVE BASE PATTERN DOCX FILE TO FOLDER OF PROGRAM FILE..")
        sys.exit()
    except ValueError as err:
        print(f"\n[-] VALUE ERROR > {str(err).upper()} <")
        input()
    except TypeError as err:
        print(f"\n[-] TYPE ERROR > {str(err).upper()} <")
        input()
    except redminelib.exceptions.AuthError:
        print(f"\n[-] INVALID USERNAME OR PASSWORD") 
        sys.exit()
    except KeyError:
        print(f"\n[-] PROGRAM DID NOT WORK BECAUSE KEY ERROR DETECTED !")
        input()
    except IndexError as err:
        print(f"\n[-] INDEX ERROR > {str(err).upper()} <")
        input()
    except redminelib.exceptions.ForbiddenError:
        print("\n[-] REQUESTED RESOURCE IS FORBIDDEN !")
        sys.exit()
    except FileNotFoundError:
        print("\n[-] SOME FILES (IMAGE, BASE DOCX) ARE NOT IN PLACE. || FILE TYPE OF IMAGES ARE MUST BE ( .png ) ")
        sys.exit()
    except docx.image.exceptions.UnrecognizedImageError:
        print("\n[-] UNRECOGNIZED IMAGE ERROR")
        sys.exit()
# #############################################################################################################
if __name__ == '__main__':
    main()
# CREATED BY © n3gat1v3o