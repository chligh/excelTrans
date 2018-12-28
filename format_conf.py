from openpyxl import load_workbook
import os, fnmatch
import codecs
import sys

class format_conf:
    def __init__(self):
        self.type = "line"
        self.out_confs = []
    
def load_format_conf(sheet):
    ret = format_conf()
    for cells in sheet.rows:

        if cells[0].value != None and type(cells[0].value) != int:
            if ".xml" in cells[0].value or ".lua" in cells[0].value or ".ini" in cells[0].value or ".cfg" in cells[0].value or ".csv" in cells[0].value or ".font" in cells[0].value or ".ts" in cells[0].value:
                ret.out_confs.append(cells[0].value)
            else:
                return ret
        else:
            return ret
    return ret

def node_desc(name):
    return name.split('.', 1)[0]
    
def is_conf_node(desc, i):
    if (1 == desc[i].value):
        return True
    return False
    
def is_keyword_node(desc, i):
    if (2 == desc[i].value):
        return True
    return False

def is_multi_key_node(desc, i):
    if (3 == desc[i].value):
        return True
    return False

def cell_value(c):
    #if (c.is_date()):
    #    return c.value.strftime("%Y-%m-%d_%H:%M:%S")
    #el
    if (c.value):
        str = unicode(c.value)
        return str.replace('|', ':')
    else:
        return "0"
        
def xml_entry(c, d):
    return u"<%s>%s</%s>"%(d, c, d)
        
def to_xml_row_str(rows, line, titleLine, row):
    return "<%s>\n%s\n</%s>"%(
        'man', 
        '\n'.join([xml_entry(cell_value(row[i]), node_desc(rows[titleLine][i].value)) for i in range(1,len(row)) if is_conf_node(rows[line], i)]), 
        'man')

def to_ini_row_str(rows, line, titleLine, row):
    return '  '.join([cell_value(row[i]) for i in range(1,len(row)) if is_conf_node(rows[line], i)])

def to_csv_row_str(rows, line, titleLine, row):
    return_value = ""
    bFirst = True
    for i in range(1,len(row)):
        if is_conf_node(rows[line], i):
            cellValue = cell_value(row[i])
            if cellValue.find(",") >= 0:
                cellValue = '"' + cellValue + '"'
            if bFirst == True:
                return_value = return_value + cellValue
                bFirst = False
            else:
                return_value = return_value + ',' + cellValue
    return return_value

def is_number(value):
    result = True
    try:
        int(value)
    except:
        result = False
    if result == False:
        try:
            result = True
            float(value)
        except:
            result = False
    return result


def to_lua_row_str(file_name, rows, line, titleLine, row):
    return_value = 'g_cfg["' + file_name + "_cfg"
    bFirst = True
    bHasMultiKey = False
    
    for i in range(1,len(row)):
        keyValue  = cell_value(rows[titleLine][i])
        cellValue = cell_value(row[i])
        
        if is_multi_key_node(rows[line], i):
            if False == bHasMultiKey:
                return_value = return_value + "$$"
                bHasMultiKey = True
            if is_number(cellValue) != True:
                cellValue = '"' + cellValue + '"'
            return_value = return_value + keyValue + "=" + cellValue + ","
        
        if is_keyword_node(rows[line], i):
            return_value = return_value + "_" + cellValue
            
        if is_conf_node(rows[line], i):
            if is_number(cellValue) != True:
                cellValue = '"' + cellValue + '"'
            
            if bFirst == True:
                return_value = return_value + '"]={' + keyValue + "=" + cellValue
                bFirst = False
            else:
                return_value = return_value + ',' + keyValue + "=" + cellValue
    return_value = return_value + "}"
    return  return_value

g_all_lua_string=""
g_all_font_string={}

def save_conf_file(filename, data):
    if data != "":
        f = codecs.open("output/" + filename, "w", "utf-8")
        f.write(data)
        f.close()

def ini_description(rows, line, titleLine):
    return '#'+'  '.join([node_desc(rows[titleLine][i].value) for i in range(1, len(rows[titleLine])) if is_conf_node(rows[line], i)])

def csv_description(rows, line, titleLine):
    return ''+','.join([node_desc(rows[titleLine][i].value) for i in range(1, len(rows[titleLine])) if is_conf_node(rows[line], i)])


def generate_conf_text(file_name, file_ext, conf_type, line, titleLine, rows):
    global g_all_lua_string
    global g_all_font_string

    totleLine = len(rows)
    for i in range(titleLine + 1, totleLine):
        if rows[i][1].value or rows[i][1].value == 0:
            continue
        else:
            totleLine = i
            break
    if (file_ext == ".xml"):
        xmlout = "<conf>\n%s\n</conf>"%('\n'.join([to_xml_row_str(rows, line, titleLine, i) for i in rows[titleLine + 1:totleLine]]))
        xmlout = u"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n" + xmlout
        return xmlout
    if (file_ext == ".ini" or file_ext == ".cfg"):
        iniout = "%s\n%s\n"%(ini_description(rows, line, titleLine), '\n'.join([to_ini_row_str(rows, line, titleLine, i) for i in rows[titleLine + 1:totleLine]]))
        return iniout
    if (file_ext == ".csv"):
        csvout = "%s\r\n%s"%(csv_description(rows, line, titleLine), '\r\n'.join([to_csv_row_str(rows, line, titleLine, i) for i in rows[titleLine + 1:totleLine]]))
        return csvout
    if (file_ext == ".lua"):
        luaout = "%s"%('\n'.join([to_lua_row_str(file_name, rows, line, titleLine, i) for i in rows[titleLine + 1:totleLine]]))
        if (g_all_lua_string != ""):
            g_all_lua_string = g_all_lua_string + "\n"
        g_all_lua_string = g_all_lua_string + luaout
        return luaout
    if (file_ext == ".font"):
        fontout = "%s\n%s"%(ini_description(rows, line, titleLine), '\n'.join([to_ini_row_str(rows, line, titleLine, i) for i in rows[titleLine + 1:totleLine]]))
        if False == g_all_font_string.has_key(file_name):
            g_all_font_string[file_name] = fontout
        else:
            g_all_font_string[file_name] = g_all_font_string[file_name] + fontout
        return ""
    return ""
        
def format_one_sheet(filename, sheet):
    fconf = None
    fconf = load_format_conf(sheet)
	
    #output conf files
    if (fconf == None or len(fconf.out_confs) <= 0):
        return
        
    # next row must be comment
    sheet_desc = sheet.rows[len(fconf.out_confs)]
    titleLine = len(fconf.out_confs);
    line = 0
    for fcode in fconf.out_confs:
        confname = fcode
        file_desc = os.path.splitext(confname)
        file_name = file_desc[0]
        file_ext  = file_desc[1]
        save_conf_file(confname, generate_conf_text(file_name, file_ext, fconf.type, line, titleLine+1, sheet.rows))
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


_DEBUG=False

if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    if _DEBUG == True:
        import pdb
        pdb.set_trace()
    g_all_lua_string = "g_cfg={}"
    g_all_font_string = {}
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
    save_conf_file("game_cfg.lua", g_all_lua_string)
    
    for name in g_all_font_string:
        file_name = name + ".font"
        save_conf_file(file_name, g_all_font_string[name])
    os.system("config.exe")
