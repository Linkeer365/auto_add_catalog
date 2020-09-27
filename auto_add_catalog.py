import win32gui
import win32con
import win32api
import win32clipboard
import os
import sys
import time

# import pyperclip

target_dir=r"D:\AllDowns\newbooks"

catalog_dir=r"D:\AllDowns\newbooks\catalogs"

ssids_path=r"D:\AllDowns\newbooks\all_ssids.txt"

ct_dir=r"D:\刺头书\ucdrs无书签"

error2_path=r"D:\AllDowns\newbooks\catalogs\error-notfetch.txt"

# strange_dir=r"D:\AllDowns\strangebooks"

pce_str="PdgCntEditor（文本）"
pce_path=r"C:\Program Files\PdgCntEditor\PdgCntEditor.exe"

pce_str2="PdgCntEditor"



def pick_good_from_ori(order,ori_text_list):
    if order=="a": # a for all
        return ori_text_list
    elif order=="d":
        return []
    elif order.isdigit():
        last_line_idx=int(order)-1
        return ori_text_list[0:last_line_idx]

def getText():
    # 读取剪切板
    # https://www.wandouip.com/t5i58809/
    win32clipboard.OpenClipboard()  # 打开剪切板

    # 使用CF_UNICODETEXT可以直接得到中文
    d = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)  # 得到剪切板上的数据

    win32clipboard.CloseClipboard()  # 关闭剪切板

    return d

def setText(aString):  # 写入剪切板
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, aString)
    win32clipboard.CloseClipboard()

def clearText():
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.CloseClipboard()

def press_ctrl_a():
    # 全选操作

    win32api.keybd_event(17,0,0,0)  #ctrl键位码是17
    win32api.keybd_event(65,0,0,0)  #A键位码是65
    win32api.keybd_event(65,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
    win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)

def press_ctrl_c():

    # 复制操作

    win32api.keybd_event(17,0,0,0)  #ctrl键位码是17
    win32api.keybd_event(67,0,0,0)  #C键位码是67
    win32api.keybd_event(67,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
    win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)

def press_ctrl_v():

    # 粘贴操作

    win32api.keybd_event(17,0,0,0)  #ctrl键位码是17
    win32api.keybd_event(86,0,0,0)  #V键位码是86
    win32api.keybd_event(86,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
    win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)

def press_ctrl_s():

    # 保存操作

    win32api.keybd_event(17,0,0,0)  #ctrl键位码是17
    win32api.keybd_event(83,0,0,0)  #s键位码是83
    win32api.keybd_event(83,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
    win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)



def click_on_pos(pos_list):
    btn_pos = pos_list
    win32api.SetCursorPos(btn_pos)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

def get_hd_from_child_hds(father_hd,some_idx,expect_name):
    child_hds=[]
    win32gui.EnumChildWindows(father_hd,lambda hwnd, param: param.append(hwnd),child_hds)

    names=[win32gui.GetWindowText(each) for each in child_hds]
    hds=[hex(each) for each in child_hds]
    print("ChildName List:",names)
    print("Child Hds List:",hds)

    name=names[some_idx]
    hd=hds[some_idx]

    print("The {} Child.".format(some_idx))
    print("The Name:{}".format(name))
    print("The HD:{}".format(hd))

    if name==expect_name:
        return child_hds[some_idx]
    else:
        print("窗口不对！")
        return None

# while True:
#     tempt = win32api.GetCursorPos() # 记录鼠标所处位置的坐标
#     print(tempt)
#     time.sleep(2)


def move_to_ctb(pdf_name):
    old_pdf_path=f"{target_dir}{os.sep}{pdf_name}"
    new_pdf_path=f"{ct_dir}{os.sep}{pdf_name}"
    os.rename(old_pdf_path,new_pdf_path)



def open_one_pdf(catalog_name,pdf_name):

    startAt=time.time()
    os.startfile(pce_path)
    # 打开文件1-2s，保险2s

    # 2020年9月27日18:42:56，优先改这里...
    time.sleep(1)

    root_hd=None
    pce_hd=win32gui.FindWindowEx(root_hd,0,0,pce_str)

    # while not pce_hd:
    #     pce_hd=win32gui.FindWindowEx(root_hd,0,0,pce_str)
        # time.sleep(1)

    dakai_pos=(414,224)
    click_on_pos(dakai_pos)

    # 使用while不大靠谱...

    time.sleep(1)

    dakai_str='打开'
    dakai_hd=win32gui.FindWindowEx(root_hd,0,0,dakai_str)

    # while not dakai_hd:
    #     click_on_pos(dakai_pos)
    #     dakai_hd=win32gui.FindWindowEx(root_hd,0,0,dakai_str)
    #     # time.sleep(1)


    # 数一下是第11个...(有三个空位去一个个试)
    feed_path_hd=get_hd_from_child_hds(dakai_hd,12,"")

    # 数一下，确定键（打开）是第16个...
    queding_hd=get_hd_from_child_hds(dakai_hd,16,"打开(&O)")

    pce_hd=win32gui.FindWindowEx(root_hd,0,0,pce_str)

    # catalog_name="10448747ssidssidtypetype1张文显《法哲学范畴研究》isbnisbn9787562020998.txt"
    catalog_path=f"{catalog_dir}{os.sep}{catalog_name}"

    # pdf_name="typetype1张文显《法哲学范畴研究》isbnisbn9787562020998.pdf"
    pdf_path=f"{target_dir}{os.sep}{pdf_name}"

    win32gui.SendMessage(feed_path_hd,win32con.WM_SETTEXT,0,pdf_path)

    # 这里send理论上是不需要sleep的...

    feed_path_hd = get_hd_from_child_hds(dakai_hd, 12, "")

    # time.sleep(2)

    win32gui.SendMessage(queding_hd,win32con.BM_CLICK)

    # 这段先留着，主要一开始我想搞缝合，所以有这段，但现在觉得其实没什么必要了...

    inside_pos=(804, 421)
    # click_on_pos(inside_pos)
    # time.sleep(1)
    #
    # press_ctrl_a()
    # press_ctrl_c()
    #
    # # 写入剪贴板需要2s...
    #
    # time.sleep(1)
    #
    # ori_text=getText()
    #
    # # clearText()
    #
    # if not ori_text:
    #     print("Empty!Empty!Empty!")
    #
    #
    # ori_text_list=ori_text.split("\n")
    #
    # for each_idx,each_line in enumerate(ori_text_list,1):
    #     print(repr(each_line),"\t\t\t",each_idx)
    # # time.sleep(5)
    # # order=input("(d for drop; a for all; num for the last line number)\nYour choice:")
    # order="d"
    # final_ori_text_list=pick_good_from_ori(order,ori_text_list)
    final_ori_text_list=[]
    catalog_list=[]
    with open(catalog_path,"r",encoding="utf-8") as f:
        catalog_list=[each.lstrip("\t").replace("\n","\r\n") for each in f.readlines() if bool(each.strip("\n"))!=0]

    final_list=final_ori_text_list+catalog_list

    # 这狗东西是用\r\n结尾的，艹！！

    for each_line in final_list:
        print(repr(each_line))

    final_catalog="".join(final_list)

    # print(repr(final_catalog))
    setText(final_catalog)

    # pyperclip.copy(final_catalog)

    # time.sleep(2)

    assert getText()==final_catalog

    win32gui.SetForegroundWindow(pce_hd)

    # 这里必须3s
    time.sleep(1)

    click_on_pos(inside_pos)
    # time.sleep(1)
    press_ctrl_a()
    press_ctrl_v()
    time.sleep(1)
    press_ctrl_s()
    time.sleep(1)

    # time.sleep(2)

    # 最后那个存储完毕你也得先点了再退出呀，做事不细心唉...

    pce2_hd=win32gui.FindWindowEx(root_hd,0,0,pce_str2)
    queding_hd=get_hd_from_child_hds(pce2_hd,0,"确定")
    win32gui.SendMessage(queding_hd,win32con.BM_CLICK)

    win32gui.SendMessage(pce_hd,win32con.WM_CLOSE,0,0)

    time.sleep(1)

    # 最后这里再把剪贴板清了...

    # clearText()

    new_name=f"withCata_{pdf_name}"
    new_path=f"{target_dir}{os.sep}{new_name}"
    os.rename(pdf_path,new_path)

    endAt=time.time()

    time_cost=endAt-startAt

    print(f"Run time:{time_cost}")
    print("one done.")

def main():
    lines=[]
    with open(ssids_path,"r",encoding="utf-8") as f1:
        lines=f1.readlines()
    lines=[each.strip("\n") for each in lines]

    with open(error2_path,"r",encoding="utf-8") as f2:
        error2s=f2.readlines()

    catalogs=os.listdir(catalog_dir)

    for each_line in lines:
        ssid,pdf_name=each_line.split("\t\t\t")
        if os.path.exists(f"{target_dir}{os.sep}withCata_{pdf_name}"):
            print("already:",pdf_name)
            continue
        pdf_name_change_suffix=pdf_name.replace(".pdf",".txt")
        expect_catalog_name=f"{ssid}ssidssid{pdf_name_change_suffix}"
        error_catalog_name=f"error_{expect_catalog_name}"
        if expect_catalog_name in catalogs and os.path.exists(f"{target_dir}{os.sep}{pdf_name}"):
            catalog_name=expect_catalog_name
            print("Catalog name:",catalog_name)
            open_one_pdf(catalog_name,pdf_name)
        else:
            if error_catalog_name+"\n" in error2s and not os.path.exists(f"{ct_dir}{os.sep}{pdf_name}"):
                move_to_ctb(pdf_name)
            else:
                print("already ctb.")
    print("all done.")

if __name__ == '__main__':
    main()












# sys.exit(0)







