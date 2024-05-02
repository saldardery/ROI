import pandas as pd
from datetime import datetime , timedelta
from dateutil.relativedelta import relativedelta

################################################################

# Read both migration jira dump excel and Mapped customer excel

################################################################


Migration_Jira_Dump = 'C:\\Users\\IncortaBI\\Box\\Incorta Backend\\Migration Jira Dump.csv'
Assessment_Jira_Dump = 'C:\\Users\\IncortaBI\\Box\\Incorta Backend\\Assessment Jira Dump.csv'
ROI_SIMPLE_VIEW='C:\\Users\\IncortaBI\\Box\\Incorta Backend\\ROI\\ROI_Simple_View.xlsx'
Migration_Jira_Dump = pd.read_csv(Migration_Jira_Dump)
ROI_SIMPLE_VIEW = pd.read_excel(ROI_SIMPLE_VIEW)
Assessment_Jira_Dump = pd.read_csv(Assessment_Jira_Dump)


################################################################

# Below code will take each Jira Name in the mapped excel and checks all completed migrations in jira dump, then it will compare all closed dates of this customer and takes the minimum closed date

################################################################


z = 0  #(it is syncing with the first for loop)

for iterate_ROI_SIMPLE_VIEW in ROI_SIMPLE_VIEW.iterrows():

        i = 0
        closed_date = []
        for iterate_migration_jira_dump in Migration_Jira_Dump.iterrows(): #
            try:
                if ((str( ROI_SIMPLE_VIEW.loc[z,'Migration Customer Name']).lower()==str(Migration_Jira_Dump.loc[i,'Custom field (Customer)']).lower())and ((Migration_Jira_Dump.loc[i,'Status']=='Closure') or (Migration_Jira_Dump.loc[i,'Status']=='Completed')))  : # compare between current Jira Mapped customer in the first for loop and current jira dump customer in the second for loop
                    date=Migration_Jira_Dump.loc[i,'Custom field (Close Date)']
                    date=datetime.strptime(date,"%d/%b/%y %I:%M %p")
                    date=date.strftime("%y/%m/%d")
                    closed_date.append(date)

                i = i + 1
            except:
                do_nohting=1
        i = 0
        for iterate_assessment_jira_dump in Assessment_Jira_Dump.iterrows():
            try:
                if ((str(ROI_SIMPLE_VIEW.loc[z,'Assessment Customer Name']).lower()==str(Assessment_Jira_Dump.loc[i,'Custom field (Customer)']).lower()) and ((Assessment_Jira_Dump.loc[i,'Status']=='Closure') or (Assessment_Jira_Dump.loc[i,'Status']=='Completed')))  : # compare between current Jira Mapped customer in the first for loop and current jira dump customer in the second for loop
                    date=Assessment_Jira_Dump.loc[i,'Custom field (Close Date)']
                    date=datetime.strptime(date,"%d/%b/%y %I:%M %p")
                    date=date.strftime("%y/%m/%d")
                    closed_date.append(date)
                i = i + 1
            except:
                do_nothing=1
        try:
            minimum_closed_date= min(closed_date)
            minimum_closed_date = datetime.strptime(minimum_closed_date, "%y/%m/%d")
            minimum_closed_date=minimum_closed_date.strftime("%d-%m-%y")
            ROI_SIMPLE_VIEW.loc[z,'First Closed Date']= minimum_closed_date

        except:
            do_nothing=1


        z = z + 1

################################################################

# Saves a new Mapped_Excel with today's date and  has "First closed date" populated to any new Missing Customers


################################################################
today_date = datetime.now().strftime("%d_%m_%Y")
#excel_filename = f"Mapped_Excel_{today_date}.csv"
#ROI_SIMPLE_VIEW.to_csv("C:\\Users\\Incorta BI\\Box\\Incorta Backend\\ROI\\ROI_simple_view.csv", index=False)
ROI_SIMPLE_VIEW.to_excel("C:\\Users\\IncortaBI\\Box\\Incorta Backend\\ROI\\ROI_Simple_View.xlsx", index=False)

