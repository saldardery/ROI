import pandas as pd
Jira_dump = 'C:\\Users\\IncortaBI\\Box\\Incorta Backend\\Migration Jira Dump.csv'
Mapped_Customer='C:\\Users\\IncortaBI\\Box\\Incorta Backend\\ROI\\ROI_Simple_View.xlsx'

Jira_dump = pd.read_csv(Jira_dump)
Mapped_Customer = pd.read_excel(Mapped_Customer)

results=[]
unique_results=[]
results_final=[]
#results= {'Missing Customer Name':[]}
#results= pd.DataFrame(results)
x=0

for Jira_dump_counter in Jira_dump.iterrows():
    flag=bool(False)
    if (Jira_dump.loc[x, 'Status'] == 'Completed' or Jira_dump.loc[x, 'Status'] == 'Closure'):

        y=0
        for Mapped_Customer_counter in Mapped_Customer.iterrows():
            if str(Mapped_Customer.loc[y,'Migration Customer Name']).lower() in str(Jira_dump.loc[x,'Custom field (Customer)']).lower():
                flag=bool(True)
                break
            y=y+1

        if(flag):
            do_nothing=1
        else:
            #results.append(Mapped_Customer.loc[x,'Jira Name'])
            results.append(Jira_dump.loc[x,'Custom field (Customer)'])


    x=x+1


for item in results:
    if item not in unique_results:
        unique_results.append(item)





test={'Missing Customer Name':unique_results}
test=pd.DataFrame(test)

test.to_csv('C:\\Users\\IncortaBI\\Box\\Incorta Backend\\ROI\\Missing Customers.csv', index=False)

