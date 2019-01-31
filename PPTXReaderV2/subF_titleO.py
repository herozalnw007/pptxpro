'''
Created on Jan 16, 2019

@author: Bowonwit
'''

class titleO():
    
    def __init__(self,io_type,io_page,io_parent,io_data):
        self.io_type = io_type
        '''  
            -1 = SLIDE
             0 = TITLE
             1 = BODY level 1
             2 = BODY level 2
               .
               .
               .
             n = BODY level n 
        '''
        self.io_page = io_page
        self.io_parent = io_parent
        self.io_data = io_data
        
        
    def set_io_type(self,io_type):
        self.io_type = io_type
    
    def set_io_page(self,io_page):
        self.io_page = io_page
    
    def set_io_parent(self,io_parent):
        self.io_parent = io_parent
    
    def set_io_data(self,io_data):
        self.io_data = io_data


    
    def get_io_type(self):
        return self.io_type
    
    def get_io_page(self):
        return self.io_page
    
    def get_io_parent(self):
        return self.io_parent
        
    def get_io_data(self):
        return self.io_data

def replace_all(text):
    dicts = [["\n\t","\t\n","\n","\t","\u2019"],
            [" "," "," "," ","'"]]
    for i in range(len(dicts[0])):
        temp = dicts[0][i]
        if i != 0 and temp == dicts[0][i]: text = text.replace(dicts[0][i], "")
        else: text = text.replace(dicts[0][i], dicts[1][i])
    tr = ""
    temp_c = True
    for t in text[::-1]:
        if t == ' ' and temp_c == True: temp_c = True
        else: temp_c = False ;tr += t
    text_return = ""
    for te in tr[::-1]:text_return += te
    return text_return

def replace_for_compare(text):
    dicts = [[" ","\n\t","\t\n","\n","\t","\u2019"],
            ["","","","","","'"]]
    for i in range(len(dicts[0])): text = text.replace(dicts[0][i], dicts[1][i])
    return text

def replace_2019(text):
    dicts = [["\u2019"],
            ["'"]]
    for i in range(len(dicts[0])): text = text.replace(dicts[0][i], dicts[1][i])
    return text

def check_has_already(listC,objC):
    for i,li in enumerate(listC):
        if li.get_io_type() == objC.get_io_type() and li.get_io_parent() == objC.get_io_parent() and replace_for_compare(li.get_io_data()) == replace_for_compare(objC.get_io_data()):
            return i
    return -1

def find_title(prs,ido_obs):
    return_obj = [
                    {
                        'object': ido_obs,
                        'pages': [0],
                    },
                ]
    for i,slide in enumerate(prs.slides):

        create_for_check = titleO(0,i+1,return_obj[0]['object'],replace_2019(slide.shapes.title.text))
        _already_index = check_has_already([ob['object'] for ob in return_obj], create_for_check)
        if _already_index < 0 :
            return_obj.append({
                                'object': create_for_check,
                                'pages': [i+1],
                            },)
        else:
            return_obj[_already_index]['pages'].append(i+1)
            create_for_check = False
    return return_obj

def tap_counter(strs):
    num = 0
    for s in strs:
        if s == '\t': num += 1
        else: return num
    return num

def revers_list(listA):
    listB = []
    for item in reversed(listA): listB.append(item)
    return listB

def last_priority_obj(ido_obs, objs, current_priority, index):
    objs_r = revers_list(objs)
    for item in objs_r[-index:(len(objs_r)+1)]:
        if item[1] == current_priority:
            for ido_ob in ido_obs:
                if replace_for_compare(item[3]) == replace_for_compare(ido_ob.get_io_data()):
                    return ido_ob.get_io_parent()
    return False

def reduce_redundance(ido_obs):
    for i,_obi in enumerate(ido_obs):
        for j,_obj in enumerate(ido_obs):
            if i != j and replace_for_compare(_obi.get_io_data()) == replace_for_compare(_obj.get_io_data()):
                for chij in find_by_parent(ido_obs, _obj):
                    chij.set_io_parent(_obi)
                ido_obs.remove(_obj)
                _obj = False

def resort_shapes(shapes):
    shape_list = []
    for shape in shapes:
        shape_list.append(shape)
    shape_list.sort(key=lambda shape: ((shape.top-shapes.title.top), (shape.left)))
    return shape_list

def is_parent(_listO, _Obj):
    for _o in _listO:
        if _o.get_io_parent() == _Obj:
            return True
    return False

def find_by_prior(_listO, parent, _pri):
    _listR = []
    for _o in _listO:
        if _o.get_io_type() == _pri and parent == _o.get_io_parent():
            _listR.append(_o)
    return _listR

def find_by_parent(_listO, parent):
    _listR = []
    for _o in _listO:
        if parent == _o.get_io_parent():
            _listR.append(_o)
    return _listR

def create_que_T_list(_listO, _listF):
    _listR = []
    _listOR = []
    for _f in _listF:
        _listR.append(str("(+) " if is_parent(_listO,_f) else "(-) ") + replace_all(_f.get_io_data()))
        _listOR.append(_f)
    return _listR,_listOR

def find_parent(ido_obs, ido_obj):
    for ido in ido_obs:
        if ido == ido_obj: return ido
    return False

def find_index_que(que_text, que_lists):
    for i,que_list in enumerate(que_lists):
        if replace_for_compare(que_list) == replace_for_compare(que_text):
            return i
    return False
            
            
            
            
            
            
        
            
