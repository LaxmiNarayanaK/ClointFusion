import ClointFusion as cf
import os
import time
import shutil

cf.OFF_semi_automatic_mode()
cf.window_show_desktop()

#copying data to completed Excel file and make changes in it
if not os.path.exists('C:/Users/Anu/clointfusion/Task-3/completed.xlsx'):
    shutil.copy('C:/Users/Anu/clointfusion/Task-3/Excel.xlsx','C:/Users/Anu/clointfusion/Task-3/completed.xlsx')

cf.launch_any_exe_bat_application("C:/Users/Anu/clointfusion/Task-3/completed.xlsx")
#To enable sort&Filter
cf.key_press('ctrl+shift+l')

#Go to Namebox
cf.key_press('alt+f3')

#Go to Region Column and filter
cf.key_write_enter("B1")
cf.key_press('alt+down')
cf.key_write_enter("ecentral")

cf.key_press('alt+f3')

#Go to Total Column
cf.key_write_enter("G1")
cf.key_press('alt+down')

#Sort Total Column in Ascending Order
cf.key_press('down')
cf.key_hit_enter()

#Green-yellow-red
cf.key_press('ctrl+space')
cf.key_press('alt+h+l+s')
cf.key_hit_enter()

#Border the Table
cf.key_press("up")
cf.key_press('ctrl+a')
cf.key_press('alt+h+b+a')

#Saving all the filters and colours
cf.key_press('ctrl+s')
cf.window_close_windows("Excel")

#Browser and outlook
cf.browser_navigate_h("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1621775318&rver=7.0.6737.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26RpsCsrfState%3da28024fb-5807-740b-f967-1e49a65bda97&id=292841&aadredir=1&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=90015")

#SIgnin outlook
cf.key_write_enter("laxminarayana33317@gmail.com")
cf.key_write_enter("Anu1234$")

time.sleep(5)
#cf.browser_locate_element_h('//*[@id="id__7"]')
cf.browser_mouse_click_h('New Message')
time.sleep(2)
cf.browser_write_h("fharookshaik.clointfusion@gmail.com",User_Visible_Text_Element="To")
cf.key_hit_enter()
cf.key_write_enter("avinash.clointfusion@gmail.com")
cf.key_write_enter("shrinidhi.clointfusion@gmail.com")



cf.browser_mouse_click_h('cc')
cf.key_write_enter("laxminarayana.clointfusion@gmail.com")

cf.browser_write_h("Task-3 Automation Test",User_Visible_Text_Element="Add a Subject")

cf.key_press('tab')
cf.key_write_enter('This is a mail sent by a bot made of ClointFusion with a dummy information.\nTable Data:')
cf.launch_any_exe_bat_application("C:/Users/Anu/clointfusion/Task-3/completed.xlsx")
cf.key_press("ctrl+a")
cf.key_press('ctrl+c')
cf.key_press('alt+f4')
cf.key_hit_enter()
cf.key_press('ctrl+v')

cf.browser_mouse_click_h("Attach")
cf.browser_mouse_click_h("Browse this computer")
cf.key_write_enter(r"C:\Users\Anu\clointfusion\Task-3\completed.xlsx")

cf.browser_mouse_click_h("Send")






















