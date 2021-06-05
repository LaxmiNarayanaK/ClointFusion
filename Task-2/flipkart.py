import ClointFusion as cf
import time
cf.OFF_semi_automatic_mode()
#Minimizes all the windows and shows the desktop
cf.window_show_desktop()
#Launches the Google Chrome
cf.launch_any_exe_bat_application("chrome") 
#Maximizes the chrome
cf.window_activate_and_maximize_windows("chrome")
#Single click on the search bar using coordinates
cf.mouse_click(950,530,single_double_triple='single')
#Types Flipkart.com and clicks Enter key
cf.key_write_enter("flipkart.com")
#Single click on the search bar in flipkart website using coordinates
cf.mouse_click(700,155,single_double_triple='single')
# add the given two products to cart by using snip return coordinates using images
cf.key_write_enter("boAt Airdopes 131 Bluetooth Headset")
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y(r"C:\Users\Anu\clointfusion\Task-2\images\airpods1.png"),single_double_triple='single')
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y(r"C:\Users\Anu\clointfusion\Task-2\images\add to cart.png"),single_double_triple='single')
cf.mouse_click(700,155,single_double_triple='single')
cf.key_write_enter("boAt Airdopes 402 Bluetooth Headset")
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y(r"C:\Users\Anu\clointfusion\Task-2\images\airpods2.png"),single_double_triple='single')
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y(r"C:\Users\Anu\clointfusion\Task-2\images\add to cart.png"),single_double_triple='single')
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y(r"C:\Users\Anu\clointfusion\Task-2\images\placeorder.png"),single_double_triple='single')
time.sleep(1.5)
cf.key_press('PageDown')
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y(r"C:\Users\Anu\clointfusion\Task-2\images\continue.png"),single_double_triple='single')
#enter CVV
cf.key_write_enter()
cf.window_close_windows("chrome")

