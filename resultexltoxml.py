import xlrd
from xml.dom import minidom

xlsx_path = 'C:\\Users\\distiction\\Desktop\\cli-view.xlsx'
xml_path = 'C:\\Users\\distiction\\\Desktop\\xmltmp\\'
xml_result = 'C:\\Users\\distiction\\\Desktop\\result\\'

def sheetnum_to_propmtstring(sheetnum):
    propmts={
    0:"enable>",
    1:"config#",
    2:"config-if-0/${ifgig}#",
    3:"config-vlan-${ifvlan}#",
    4:"if-${ifgig} switchport access#",
    5:"if-${ifgig} switchport trunk#",
    6:"acl-ingress-${aclid}#",
    7:"acl-egress-${aclid}#",
    8:"routeStatic#",
    9:"NAT#",
    10:"dhcpserver#"
    }
    return propmts.get(sheetnum,"#")

def padexcelcell(mulcommands,row):
    commands=[]
    if mulcommands[0][0] > 2:
        firstrow = 2
        for singlenum in range(0,mulcommands[0][0]-2):
            if len(table.cell_value(firstrow,mulcommands[0][2])) != 0:
                singlecommand = (firstrow,firstrow+1,mulcommands[0][2],mulcommands[0][3])
                firstrow+=1
                commands.append(singlecommand)
    for i,mulcommand in enumerate(mulcommands):
        if mulcommand == mulcommands[-1]:
            singlenum = row - mulcommands[-1][1]
        else:
            singlenum = mulcommands[i+1][0]-mulcommands[i][1]
            #print(singlenum,mulcommands[i][1],mulcommands[i+1][0])
        commands.append(mulcommand)
        if singlenum > 0 :
            singlecommandrow = mulcommands[i][1]
            #print(singlecommandrow)
            for single in range(singlenum):
                singlecommand = (singlecommandrow+single,singlecommandrow+single+1,mulcommand[2],mulcommand[3])
                commands.append(singlecommand)
    return commands

def paramcreate(fatherElement,table,fatherlastrow,row,col): 
    currentrow = row
    while currentrow < fatherlastrow: #name
        while len(table.cell_value(currentrow,col)) != 0 and currentrow < fatherlastrow :
            ####添加节点
            paramElement = dom.createElement('PARAM')
            #print(table.cell_value(currentrow,col)) #name
            paramElement.setAttribute('name',table.cell_value(currentrow,col))
            if len(table.cell_value(currentrow,col+1)) !=0 : #test
                paramElement.setAttribute('test',str(table.cell_value(currentrow,col+1)))
            #help不让为空
            paramElement.setAttribute('help',table.cell_value(currentrow,col+2))
            #order
            if table.cell_value(currentrow,col+3):
                paramElement.setAttribute('order',str(bool(table.cell_value(currentrow,col+3))))
            if table.cell_value(currentrow,col+4): 
                paramElement.setAttribute('optional',str(bool(table.cell_value(currentrow,col+4))))
            if table.cell_value(currentrow,col+5): 
                paramElement.setAttribute('mode',table.cell_value(currentrow,col+5))
            paramElement.setAttribute('ptype',table.cell_value(currentrow,col+6))
            fatherElement.appendChild(paramElement)
            ###添加节点
            fatherstartrow = currentrow
            currentrow +=1
            #if len(table.cell_value(currentrow,col)) == 0 and currentrow < fatherlastrow:
            #查看本节点是否跨行，跨行取下一个节点
            if currentrow == fatherlastrow:
                if table.cell_value(fatherstartrow,col+7) != 0:
                    if isinstance(table.cell_value(fatherstartrow,col+7),int):         
                        paramcreate(paramElement,table,currentrow,fatherstartrow,col+8)
                        fatherstartrow = currentrow
                break
            try:
                while len(table.cell_value(currentrow,col)) == 0 and currentrow < fatherlastrow :
                    currentrow +=1
                #超过范围生成子参数    
                if (len(table.cell_value(currentrow,col)) != 0 or currentrow == fatherlastrow) and table.cell_value(fatherstartrow,col+7) != 0:
                    if isinstance(table.cell_value(fatherstartrow,col+7),int):         
                        paramcreate(paramElement,table,currentrow,fatherstartrow,col+8)    
            except:
                if table.cell_value(fatherstartrow,col+7) != 0:
                    if isinstance(table.cell_value(fatherstartrow,col+7),int):
                        paramcreate(paramElement,table,currentrow,fatherstartrow,col+8)
                break
            finally:
                fatherstartrow = currentrow
            """ while len(table.cell_value(currentrow,col)) == 0 and currentrow < fatherlastrow :
                currentrow +=1
                try:
                    table.cell_value(currentrow,col)
                except:
                    if table.cell_value(fatherstartrow,col+6) != 0:
                        if isinstance(table.cell_value(fatherstartrow,col+6),int):
                            paramcreate(paramElement,table,currentrow,fatherstartrow,col+7)
                    break
            #超过范围生成子参数    
            if len(table.cell_value(currentrow,col)) != 0 and table.cell_value(fatherstartrow,col+6) != 0:
                if isinstance(table.cell_value(fatherstartrow,col+6),int):         
                    paramcreate(paramElement,table,currentrow,fatherstartrow,col+7)
            fatherstartrow = currentrow """
                
#excel操作
xmlexcel = xlrd.open_workbook(xlsx_path)
for sheetnum in range(0,8):#len(data.sheets())
    table = xmlexcel.sheets()[sheetnum]
    # 获取表格行数nrows
    nrows = table.nrows
    #print("表格一共有",nrows,"行")
    #获取列表的有效列数
    ncols = table.ncols   
    #print("表格一共有",ncols,"列")

    #xml操作
    dom = minidom.getDOMImplementation().createDocument(None,'CLISH_MODULE',None)
    clish_module = dom.documentElement
    clish_module.setAttribute('xmlns','http://clish.sourceforge.net/XMLSchema')
    clish_module.setAttribute('xmlns:xsi','http://www.w3.org/2001/XMLSchema-instance')
    clish_module.setAttribute('xsi:schemaLocation','http://clish.sourceforge.net/XMLSchema\n\
            http://clish.sourceforge.net/XMLSchema/clish.xsd') 

    view = dom.createElement('VIEW')
    view.setAttribute('name',table.name+'-view')
    view.setAttribute('prompt',sheetnum_to_propmtstring(sheetnum))

    #合并格提取command
    merged = table.merged_cells
    merged.sort( key =lambda x:(x[0],x[2]) )
    #print(merged)
    mulcommands = [list for list in merged if list[2]==1]
    #print(mulcommands)
    if not mulcommands:
            commands=[]
            for row in range(2,len(table.col(1))):
                commands.append((row,row+1,1,2))
    else:
        commands = padexcelcell(mulcommands,nrows)   
    #print(commands)
    #mulparams = [list for list in merged if list[2]==6]
    #print(mulparams)
    #params = padexcelcell(mulparams,nrows)



    for command in commands:
        row = command[0]
        col = command[2]

        cell_value = table.cell_value(row,col)

        """ 
        print(command)
        print( table.cell_value(row,col) ) # COMMAND
        print( table.cell_value(row,col+1) ) #help
        print( table.cell_value(row,col+2) ) #lock
        print( table.cell_value(row,col+3) ) #interrupt 
        """

        comElement = dom.createElement('COMMAND')
        comElement.setAttribute('name',table.cell_value(row,col))
        comElement.setAttribute('help',table.cell_value(row,col+1))
        comElement.setAttribute('lock',str(bool(table.cell_value(row,col+2))))
        comElement.setAttribute('interrupt',str(bool(table.cell_value(row,col+3))))
        if len(table.cell_value(row,col+4)) !=0 :
            comElement.setAttribute('view',table.cell_value(row,col+4))
        if len(table.cell_value(row,col+5)) !=0 :
            comElement.setAttribute('viewid',table.cell_value(row,col+5))

        if isinstance(table.cell_value(row,col+6),int) and table.cell_value(row,col+6) ==1 :#param
            """for prow in range(command[0],command[1]):
                if len(table.cell_value(prow,col+5)) ==0 :
                    break"""
            paramcreate(comElement,table,command[1],command[0],col+7)
        actelement = dom.createElement('ACTION')
        actelement.appendChild(dom.createTextNode(table.cell_value(command[0],ncols-1)))
        comElement.appendChild(actelement)
        view.appendChild(comElement)
        clish_module.appendChild(view)

    print(xml_path+table.name+".xml")    
    with open(xml_path+table.name+".xml", 'w+', encoding='utf-8') as f:
        dom.writexml(f, indent='\n',addindent='\t', newl='\n',encoding='utf-8')
        #pattern = re.compile('')
        #print(type(f))
    with open(xml_path+table.name+".xml", 'r+', encoding='utf-8') as f:
        xml_temp = f.read()
        xml_data = xml_temp.replace('test="&quot;',"test='\"").replace("&quot;\"",'"\'').replace("&quot;",'"').replace('</ACTION>','\t\t\t</ACTION>').replace(r'\r\n',r'\n')
        xml_data = xml_data.replace("swcli","echo swcli").replace("clicfg","echo clicfg")
        xmldatabytes = xml_data.encode()
    with open(xml_result+table.name+".xml",'wb') as f:
        f.write(xmldatabytes)