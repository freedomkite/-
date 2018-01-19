#encoding:utf-8
import xlrd
import xlsxwriter
import xlwt
'''首先读取所要处理的文件，其中源文件分为两个，一个是单义词，'''
def readfirst(src1):
    data=xlrd.open_workbook(src1)
    table=data.sheets()[2]
    nrows=table.nrows
    ncols=table.ncols
    #此处设置一个字典，用来存储读取的文件内容，其中以词语为键值，以词语的具体内容为value
    src1_dict={}
    for i in range(1,nrows):
        if table.cell(i,4).value not in src1_dict:
            src1_dict[table.cell(i,4).value]=[]
            src1_dict[table.cell(i,4).value].append(table.row_values(i)[4:7]+[table.cell(i,13).value])
        else:
            src1_dict[table.cell(i,4).value].append(table.row_values(i)[4:7]+[table.cell(i,13).value]) 
    return src1_dict

'''其次读取所要处理的第二个文件，'''
def readsecond(src2):
    data=xlrd.open_workbook(src2)
    table=data.sheets()[0]
    nrows=table.nrows
    ncols=table.ncols
    src2_dict={}
    word_list=[]
    s0='0'
    s1='1'
    s2='2'
    for i in range(3,nrows):
        if table.cell(i,1).value:
            #print table.cell(i,1).value
            tmp_word=table.cell(i,1).value   #词语
            word_list.append(tmp_word)
            tmp_pinyin=table.cell(i,2).value   #拼音
            tmp_bianma=table.cell(i,3).value   #编码
            tmp_shiyi=table.cell(i,4).value    #释义
            tmp_shili=table.cell(i,5).value    #示例
            src2_dict[table.cell(i,1).value]={}
            src2_dict[table.cell(i,1).value][s0]=[]
            src2_dict[table.cell(i,1).value][s1]=[]
            src2_dict[table.cell(i,1).value][s2]=[]
            
            src2_dict[table.cell(i,1).value][s0].append(tmp_word)
            src2_dict[table.cell(i,1).value][s0].append(tmp_pinyin)
            src2_dict[table.cell(i,1).value][s0].append(tmp_bianma)
            src2_dict[table.cell(i,1).value][s0].append(tmp_shiyi)
            src2_dict[table.cell(i,1).value][s0].append(tmp_shili)
            
            if table.cell(i,6).value:
                src2_dict[table.cell(i,1).value][s1].append(table.row_values(i)[6:])
            elif table.cell(i,12).value:
                src2_dict[table.cell(i,1).value][s2].append(table.row_values(i)[6:])  
        else:
            if table.cell(i,6).value:
                src2_dict[tmp_word][s1].append(table.row_values(i)[6:])
            elif table.cell(i,12):
                src2_dict[tmp_word][s2].append(table.row_values(i)[6:])
    return src2_dict,word_list
#   [%当事 离 2017年 1月 20日 正式 宣誓 %] [# 上任 #] [%内容 还有 2 个 月 的 时间 %] ， 
def preprocess(sentence):
    buff=[]
    s_dict={}
    if '[%' not in sentence:
        return buff
    else:
        tmp=sentence.split('%]')
        for w in tmp:		
            if '[#' in w:
				if '[%' in w:
					if w.index('[#')<w.index('[%'):
						buff.append('pred')
						ind1=w.index('[%')
						s1=w[ind1+2:]
						ind2=s1.index(' ')
						s2=s1[:ind2]
						buff.append(s2)
						s_dict[s2]=s1[ind2+1:]
				else:
					buff.append('pred')
            else:
				if '[%' in w:
					ind1=w.index('[%')
					s1=w[ind1+2:]
					ind2=s1.index(' ')
					s2=s1[:ind2]
					buff.append(s2)
					s_dict[s2]=s1[ind2+1:]
				else:
					pass
        return buff,s_dict
#总处理程序
def process(src1_xls,src2_xls,res_xls):
    s0='0'  #表示前面共有的项
    s1='1'  #表示基本语义框架结构
    s2='2'  #表示扩展语义框架结构
    #wb=xlsxwriter.Workbook(res_xlsx) # 建立文件
    #ws=wb.add_worksheet('sheet1')
    wt = xlwt.Workbook()
    ws=wt.add_sheet('Sheet1')
    src1_dict=readfirst(src1_xls)
    src2_dict,word_list=readsecond(src2_xls)
    num=3   #总的数目
    num2=0 
    num_word=0

    sent=[u'施事',u'同事',u'当事',u'接事 ',u'受事',
          u'系事',u'与事',u'结果',u'对象',u'内容',
          u'工具',u'材料',u'方式',u'原因',u'目的',
          u'事量',u'空间 ',u'时间',u'范围',u'起点',
          u'终点',u'路径',u'方向',u'处所',u'起始',
          u'结束',u'时点',u'时段']
    flag=True
    f_w=open(u'例句统计信息.txt','w')
    f_w.write((u'编码'+'\t'+u'词语'+'\t'+u'合并前'+'\t'+u'合并后'+'\n').encode('utf-8'))
    #for word in src2_dict:
    for word in word_list:
        num_word+=1
        if word in src1_dict:
            sentences=src1_dict[word]
            f_w.write((str(num_word)+'\t'+word+'\t'+str(len(src2_dict[word][s1])+len(src2_dict[word][s2]))+'\t').encode('utf-8'))
            for sentence in sentences:
                buff,s_dict=preprocess(sentence[-1])
                for w in buff:
                    if w  in sent[10:]:
                        flag=False
                        break
                if flag==False:
                    liju=sentence[1]+':'+sentence[-1]
                    laiyuan='new'
                    is_type=''
                    struct=''.join(['['+w+']' for w in buff])
                    type_struct=''
                    is_type_struct=''
                    ibuff=[' ']*28
                    for key in s_dict:
                        if key in sent:
                            ibuff[sent.index(key)]=s_dict[key]
                    obuff=['  ']*6
                    obuff.append(liju)
                    obuff.append(laiyuan)
                    obuff.append(is_type)
                    obuff.append(struct)
                    obuff.append(type_struct)
                    obuff.append(is_type_struct)
                    obuff+=ibuff
                    src2_dict[word][s2].append(obuff)
                    flag=True
                else:
                    liju=sentence[1]+':'+sentence[-1]
                    laiyuan='new'
                    is_type=''
                    struct=''.join(['['+w+']' for w in buff])
                    type_struct=''
                    is_type_struct=''
                    ibuff=[' ']*28
                    for key in s_dict:
                        if key in sent:
                            ibuff[sent.index(key)]=s_dict[key]
                    obuff=[]
                    obuff.append(liju)
                    obuff.append(laiyuan)
                    obuff.append(is_type)
                    obuff.append(struct)
                    obuff.append(type_struct)
                    obuff.append(is_type_struct)
                    obuff+=['  ']*6
                    obuff+=ibuff
                    src2_dict[word][s1].append(obuff)
            if len(src2_dict[word][s1])>0:
                for w in src2_dict[word][s1]:
                    for i,w_i in enumerate(w):
                        #print len(w), word,i+6,w_i,num+num2
                        ws.write(num+num2,i+6,w_i)
                    num2+=1
            if len(src2_dict[word][s2])>0:
                for w in src2_dict[word][s2]:
                    for i,w_i in enumerate(w):
                        ws.write(num+num2,i+6,w_i)
                    num2+=1
            
            if num2!=0:
				ws.write_merge(num,num+num2-1,0,0,num_word)
				for j,w in enumerate(src2_dict[word][s0]):
					ws.write_merge(num,num+num2-1,j+1,j+1,w) 
            else:
				ws.write_merge(num,num+num2,0,0,num_word)
				for j,w in enumerate(src2_dict[word][s0]):
					ws.write_merge(num,num+num2,j+1,j+1,w) 
				num2+=1
            num+=num2
            num2=0
            f_w.write((str(len(src2_dict[word][s1])+len(src2_dict[word][s2]))+'\n').encode('utf-8'))
            #print num
        else:
            f_w.write((str(num_word)+'\t'+word+'\t'+str(len(src2_dict[word][s1])+len(src2_dict[word][s2]))+'\t'+str(len(src2_dict[word][s1])+len(src2_dict[word][s2]))+'\n').encode('utf-8'))
            if len(src2_dict[word][s1])>0:
                for w in src2_dict[word][s1]:
                    for i,w_i in enumerate(w):
                        ws.write(num+num2,i+6,w_i)
                    num2+=1
            if len(src2_dict[word][s2])>0:
                for w in src2_dict[word][s2]:
                    for i,w_i in enumerate(w):
                        ws.write(num+num2,i+6,w_i)
                    num2+=1
            
            if num2!=0:
				ws.write_merge(num,num+num2-1,0,0,num_word)
				for j,w in enumerate(src2_dict[word][s0]):
					ws.write_merge(num,num+num2-1,j+1,j+1,w) 
            else:
				ws.write_merge(num,num+num2,0,0,num_word)
				for j,w in enumerate(src2_dict[word][s0]):
					ws.write_merge(num,num+num2,j+1,j+1,w) 
				num2+=1
            num+=num2
            num2=0
            pass
    #wb.close()
    wt.save(res_xls)
    f_w.close()

if __name__=='__main__':
    import sys
    process(sys.argv[1].decode('utf-8'),sys.argv[2].decode('utf-8'),sys.argv[3].decode('utf-8'))
    #process(u'语料筛选处理结果.xls',u'results（单义词_二校 一万句）_已选定例句 及框架20170911.xls',u'sum.xls')
    

    

                
    
        
        
        
        
            
