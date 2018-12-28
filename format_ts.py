from openpyxl import load_workbook
import os, fnmatch
import codecs
import sys

g_ts_str = ""
g_ts_files = []
   
def load_format_conf(sheet):
    out_confs = []
    for cells in sheet.rows:
        if cells[0].value != None and type(cells[0].value) != int:
            if ".xml" in cells[0].value or ".lua" in cells[0].value or ".ini" in cells[0].value or ".cfg" in cells[0].value or ".csv" in cells[0].value or ".font" in cells[0].value or ".ts" in cells[0].value:
                out_confs.append(cells[0].value)
            else:
                return out_confs
        else:
            return out_confs
    return out_confs

def node_desc(name):
    return name.split('.', 1)[0]

def cell_value(c):
    #if (c.is_date()):
    #    return c.value.strftime("%Y-%m-%d_%H:%M:%S")
    #el
    if (c.value):
        str = unicode(c.value)
        return str.replace('|', ':')
    else:
        return "0"

def save_conf_file(filename, data):
    if data != "":
        f = codecs.open("output/" + filename, "w", "utf-8")
        f.write(data)
        f.close()

def to_ts_cell_str(cname, ctype, cvalue):
    s = cname + ": "
    if ctype == "string":
        if cvalue == None:
            cvalue = ""
        s = s + "\"" + cvalue + "\""
    elif ctype == "number[]":
        if cvalue == None:
            cvalue = ""
        s = s + "[" + str(cvalue) + "]"
    elif ctype == "number":
        if cvalue == None:
            cvalue = 0
        s = s + str(cvalue)
    else:
        assert(False)
    return s

def to_ts_row_str(row, fields, index, types):
    s = "    [" + str(row[index[0]].value) + "]: { " + ", ".join([to_ts_cell_str(fields[i], types[i], row[index[i]].value) for i in range(0, len(fields))]) + " }"
    return s

def generate_conf_text(file_name, file_ext, typeLine, titleLine, rows):
    global g_ts_files
    global g_ts_str
    g_ts_files.append(file_name)
    totalLine = len(rows)
    for i in range(titleLine + 1, totalLine):
        if rows[i][1].value or rows[i][1].value == 0:
            continue
        else:
            totalLine = i
            break
    # 
    fields, index, types = [], [], []
    for i in range(1, len(rows[typeLine])):
        t = rows[typeLine][i].value
        if t != None:
            assert(t == "string" or t == "number" or t == "number[]")
            fields.append(rows[titleLine][i].value)
            index.append(i)
            types.append(t)
    if file_ext == ".ts":
        s = "const " + file_name + ": { [key: " + "number" + "]: " + "{" 
        s = s + ",".join([(" " + fields[i] + ": " + types[i]) for i in range(0, len(fields))])
        s = s + " } } = {\r\n"
        s = s + ",\r\n".join([to_ts_row_str(row, fields, index, types) for row in rows[titleLine + 1: totalLine]])
        s = s + "\r\n}\r\n\r\n"
        g_ts_str = g_ts_str + s
        return ""
    return ""
        
def format_one_sheet(filename, sheet):
    out_confs = load_format_conf(sheet)
	
    #output conf files
    if (len(out_confs) <= 0):
        return
        
    # next row must be comment
    titleLine = len(out_confs)
    line = 0
    for fcode in out_confs:
        confname = fcode
        file_desc = os.path.splitext(confname)
        file_name = file_desc[0]
        file_ext  = file_desc[1]
        if file_ext == ".ts":
            save_conf_file(confname, generate_conf_text(file_name, file_ext, line, titleLine+1, sheet.rows))
        line = line + 1

def format_one_conf(filename):
    if (filename.startswith('./~')):
        return
    print "now before process xlsm file %s"%(filename)
    wb = load_workbook(filename = filename)
    print "now process xlsm file %s"%(filename)

    for i in wb.get_sheet_names():
        if False == ("$$" in i):
            print "now process %s" %(i)
            format_one_sheet(filename, wb.get_sheet_by_name(i))

if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    #import pdb
    #pdb.set_trace()
    rootpath = "./"
    excel_pat = "*.xlsm"
    if False == os.path.exists("output") :
        os.mkdir("output")
    for root,dirs,files in os.walk(rootpath):
        for filespath in files:
            if (rootpath != root):
                continue
            
            if (not fnmatch.fnmatch(filespath, excel_pat)):
                continue
                
            fullname = os.path.join(root, filespath)
            format_one_conf(fullname)
    
    g_ts_str = g_ts_str + "const Configs = {\r\n" + ",\r\n".join(["    " + file + ": " + file for file in g_ts_files ]) + "\r\n}\r\n\r\nexport default Configs;\r\n"
    save_conf_file("Configs.ts", g_ts_str)
