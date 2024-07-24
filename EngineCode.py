with open("Services.txt","r") as file:
    services_data=file.readlines()

def add_nat_obj(word,nat_line_number):
    for i,services_line in enumerate(services_data):
        if(i==nat_line_number):
            continue
        if word in services_line.split():
            end_line = i+1
            while end_line < len(services_data) and services_data[end_line].startswith(" "):
                end_line += 1
            config_details = "".join(services_data[i:end_line]).strip()    
            config_list.append(config_details)
def extract_segment(start_line):

    #Extracting segment
    end_line = start_line
    while start_line >= 0 and  services_data[start_line].startswith(" "):
        start_line -= 1
    while end_line < len(services_data) and services_data[end_line].startswith(" "):
        end_line += 1
    config_details = "".join(services_data[start_line:end_line]).strip()    
    config_list.append(config_details)

    #Searching Using Last word of each object network line
    last_word=services_data[start_line].split()[-1]
    for i,services_line in enumerate(services_data):
        if(i==start_line):
            continue
        if last_word in services_line.split():
            config_list.append(services_line.rstrip('\n'))
            words=services_line.split()
            if words[0]=="nat":
                for word in reversed(words):
                    if "pub" in word.lower():
                        add_nat_obj(word,i)



import pandas as pd

df=pd.read_excel("Sheet.xlsx")
ips=df['GDC Private IP']

ip_addresses=df['GDC Private IP'].str.strip().dropna()
config_list=[]

for ip in ips:
    for i,services_line in enumerate(services_data):
        if ip in services_line.split():
            if(services_line.startswith(" ")):
                extract_segment(i)
            else:
                config_list.append(services_line.rstrip("\n"))
    config_list.append("\n")
    config_list.append("-----------------------X------------------------")
    config_list.append("\n")
output="\n".join(config_list)
with open("Output.txt","w") as file:
    file.write(output)
    