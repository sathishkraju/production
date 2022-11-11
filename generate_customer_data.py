import requests
import json
import xlsxwriter
from datetime import datetime
import git
from git import Repo
import sys
sys.path.append('../config')
import config

partnerId=config.partnerId
clientId=config.clientId
skey=config.skey
sval=config.sval
repo_dir=config.repo_dir
gtoken=config.gtoken

git_url=config.git_url

epurl=config.epurl
aurl = config.aurl

git_url=config.git_url

now = datetime.now()
ctime = now.strftime("%m-%d-%Y-%H-%M-%S")
# Cretae a xlsx file



apayload='grant_type=client_credentials&client_id='+skey+'&client_secret='+sval
aheaders = {
  'Content-Type': 'application/x-www-form-urlencoded',
  'Accept': 'application/json',
}

aresponse = requests.request("POST", aurl, headers=aheaders, data=apayload)
aresp=json.loads(aresponse.text)
atoken=aresp["access_token"];
bearer="Bearer "+atoken

url = epurl+"/api/v2/tenants/"+clientId+"/resources/search?queryString=type:DEVICE&state:active"

payload = "{subject:TestSubject,description:TestDescription,priority:Very Low}"
headers = {
  'Authorization': bearer,
  'Content-Type': 'application/json',
  'Accept': 'application/json'
}

response = requests.request("GET", url, headers=headers, data=payload)
res=response.json().get("results")
fname=str(res[0]["client"]["name"])+str("-")+str(ctime)
com_file=repo_dir+"/"+str(res[0]["client"]["name"])+'.xlsx';
xlsx_File = xlsxwriter.Workbook(com_file)
bold = xlsx_File.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'color':'white',
    'fg_color': '#1e4f87'})
merge_format = xlsx_File.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'color':'white',
    'fg_color': '#879b20'})
merge_format1 = xlsx_File.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'color':'white',
    'fg_color': '#002855'})
merge_format2 = xlsx_File.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'color':'#1e4f87',
    'fg_color': '#f6ea06'})
j=0
k=0
chckarr=[]
newarr={}

for i in res:
    if "make" not in res[j]:
        res[j]["make"]="-"
    if "model" not in res[j]:
        res[j]["model"]="-"
    if "source" not in res[j]:
        res[j]["source"]="-"  
    if "deviceType" not in res[j]:
        res[j]["deviceType"]="-" 
    if "hostName" not in res[j]:
        res[j]["hostName"]="-" 
    if "osName" not in res[j]:
        res[j]["osName"]="-" 
    if "biosName" not in res[j]:
        res[j]["biosName"]="-" 
    if "biosVersion" not in res[j]:
        res[j]["biosVersion"]="-" 
    if "manufacturer" not in res[j]:
        res[j]["manufacturer"]="-" 
    if "description" not in res[j]:
        res[j]["description"]="-"   
    if "agentVersion" not in res[j]:
        res[j]["agentVersion"]="-"
    if "resourceType" not in res[j]:
        res[j]["resourceType"]="-"
    if "serialNumber" not in res[j]:
        res[j]["serialNumber"]="-" 
    if "location" not in res[j]:
        location="-" 
    else :
        if str(res[j]["location"]["city"])!= "" and str(res[j]["location"]["name"])=="" :
            location=str(res[j]["location"]["city"])+","+str(res[j]["location"]["name"])
        elif str(res[j]["location"]["name"])!= "" :
            location=str(res[j]["location"]["name"])
        elif str(res[j]["location"]["city"])!= "" :
            location=str(res[j]["location"]["city"])
        else:
            location="-"        
    #print(str(res[j]["devicePath"])+"=============="+str(res[j]["deviceType"])+"=============="+str(res[j]["status"])+"=============="+str(res[j]["state"]))
    devtype=str(res[j]["devicePath"]) 
    devst=str(res[j]["status"])
    devrt=str(res[j]["resourceType"])  
    keyval=str(devtype)
    if "Network Device" in res[j]["devicePath"]:
        keyval="Network Device"

    if keyval == "Network Device":
        url1 = str(epurl+"/api/v2/tenants/"+clientId+"/resources/")+str(res[j]["id"])
        response1 = requests.request("GET", url1, headers=headers, data=payload)
        res_add=response1.json()
        if "generalInfo" not in res_add:
            firmware_version="-"
            sofwtare_version="-"
        else:
            if "firmwareVersion" in res_add["generalInfo"]:
                firmware_version=res_add["generalInfo"]["firmwareVersion"]               
            else:
                firmware_version="-"
            if "softwareVersion" in res_add["generalInfo"]:
                software_version=res_add["generalInfo"]["softwareVersion"]               
            else:
                software_version="-"               
                
        if ">>" in res[j]["devicePath"]:
            dtmain = devtype.split(" >> ")
            keyval=str(dtmain[0])
            res[j]["parent"]=keyval
            res[j]["child"]=str(res[j]["resourceType"])
            if keyval not in newarr:
                newarr[keyval]={}            
            if devrt not in newarr[keyval]:
                newarr[keyval][devrt]={}   
            len_index=len(newarr[keyval][devrt])
            newarr[keyval][devrt][len_index]=[location,str(res[j]["name"]),str(res[j]["ipAddress"]),str(res[j]["deviceType"]),str(res[j]["serialNumber"]),str(res[j]["make"]),str(res[j]["model"]),str(res[j]["biosName"]),str(res[j]["biosVersion"]),str(res[j]["manufacturer"]),str(res[j]["osName"]),str(res[j]["description"]),str(res[j]["agentVersion"]),str(res[j]["state"]),str(res[j]["status"]),str(res[j]["source"]),str(res[j]["devicePath"]),str(firmware_version),str(software_version)]        
            #print("--"+keyval+"--")
            chckarr.append(keyval)
        else:
            #print("--"+keyval+"--")
            res[j]["parent"]=str(res[j]["resourceType"])
            res[j]["child"]=str(res[j]["resourceType"])
            if keyval not in newarr:
                newarr[keyval]={}
            len_index1=len(newarr[keyval])
            newarr[keyval][len_index1]=[location,str(res[j]["name"]),str(res[j]["ipAddress"]),str(res[j]["deviceType"]),str(res[j]["serialNumber"]),str(res[j]["make"]),str(res[j]["model"]),str(res[j]["biosName"]),str(res[j]["biosVersion"]),str(res[j]["manufacturer"]),str(res[j]["osName"]),str(res[j]["description"]),str(res[j]["agentVersion"]),str(res[j]["state"]),str(res[j]["status"]),str(res[j]["source"]),str(res[j]["devicePath"]),str(firmware_version),str(software_version)]        
            chckarr.append(keyval)
    if str(res[j]["classCode"])=="vmwarehost":
        keyval="VMware"
    if keyval == "VMware":
        url1 = str(epurl+"/api/v2/tenants/"+clientId+"/resources/")+str(res[j]["id"])
        response1 = requests.request("GET", url1, headers=headers, data=payload)
        res_add=response1.json()
        firmware_version="-"
        software_version="-"
        
        if res_add["cpus"]:
            pname = res_add["cpus"][0]["processorName"].split("@")
            #print(res_add["cpus"][0]["processorName"]+"----"+pname[1]+"----")
            software_version=pname[1]             
                
        if ">>" in res[j]["devicePath"]:
            dtmain = devtype.split(" >> ")
            keyval=str(dtmain[0])
            res[j]["parent"]=keyval
            res[j]["child"]=str(res[j]["resourceType"])
            if keyval not in newarr:
                newarr[keyval]={}            
            if devrt not in newarr[keyval]:
                newarr[keyval][devrt]={}   
            len_index=len(newarr[keyval][devrt])
            newarr[keyval][devrt][len_index]=[location,str(res[j]["name"]),str(res[j]["model"]),str(res[j]["ipAddress"]),str(res[j]["osName"]),"-","-","-",str(software_version),"-"]       
            #print("--"+keyval+"--")
            chckarr.append(keyval)
        else:
            #print("--"+keyval+"--")
            res[j]["parent"]=str(res[j]["resourceType"])
            res[j]["child"]=str(res[j]["resourceType"])
            if keyval not in newarr:
                newarr[keyval]={}
            len_index1=len(newarr[keyval])
            newarr[keyval][len_index1]=[location,str(res[j]["name"]),str(res[j]["model"]),str(res[j]["ipAddress"]),str(res[j]["osName"]),"-","-","-",str(software_version),"-"]        
            chckarr.append(keyval)
    k=k+1
    j=j+1
result = [*set(chckarr)]
json_data = json.dumps(newarr)
resp = json.loads(json_data)

for r in resp:
    # Add new worksheet
    m=2
    inc=0
    nd_array={}
    if r == "Network Device":
        nd_array=["Location","Name","ipAddress","Device Type","Serial Number","Make","Model","Bios Name","Bios Version","Manufacturer","OS Name","Description","agent Version","State","Status","Source","Device Path","Firmware Version","Firmware Software"]
        sp="A"
        ep="S"
    if r == "VMware":
        nd_array=["Location","Device Name","Device Model","IP Address","Current VMware Version","Recommended Version","HyperFlex Version (if applicable)","HyperFlex Version Recommended","Server Hardware Firmware Version","Recommended Firmware"]  
        sp="A"
        ep="J"
    nd_len=len(nd_array)
    m=m+1
    sheet_schedule = xlsx_File.add_worksheet(str(r))
    sheet_schedule.merge_range(str(sp)+"1"+":"+str(ep)+"1",str(res[0]["client"]["name"]), merge_format2)
    n=nd_len
    for i in range(n):
        sheet_schedule.write(1,i,str(nd_array[i]),bold)
        
    dyncprow=str(sp)+str(m)+":"+str(ep)+str(m); 
    for y in resp[r]:        
        incc=0
        if(len(y) < 4):
            #print(resp[r][y])
            for x1 in range( nd_len):
                sheet_schedule.write(m,x1,str(resp[r][y][x1]))          
            inc=inc+1
            m=m+1
        else:      
            subc = json.dumps(resp[r][y])
            subc1 = json.loads(subc)
            m=m+1
            dyncrow=str(sp)+str(m)+":"+str(ep)+str(m)+"";             
            for z in subc1:               
                for x2 in range(nd_len):
                    sheet_schedule.write(m,x2,str(resp[r][y][z][x2]))
                m=m+1
                inc=inc+1
                incc=incc+1
            sheet_schedule.merge_range(dyncrow,str(y)+"("+str(incc)+")", merge_format1)
    sheet_schedule.merge_range(dyncprow,str(r)+"("+str(inc)+")", merge_format)
# Close the Excel file
xlsx_File.close()

repo = git.Repo(repo_dir)

# Provide a list of the files to stage
repo.index.add(str(res[0]["client"]["name"])+str('.xlsx'))
# Provide a commit message
repo.index.commit(str('Generated the file for Customer')+str(res[0]["client"]["name"])+" in the production server")
origin = repo.remote('origin')
origin.push()