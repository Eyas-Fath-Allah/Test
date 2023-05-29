import pyautogui
import time
import cv2
from openpyxl import workbook, load_workbook

# Open Excel file and get Protocol numbers
load_excel_path = ''
wb = load_workbook(load_excel_path)
ws = wb.active


# # Click to Avicenna
# pyautogui.click(30, 10)
# time.sleep(0.5)
#
# # Write Patient No.
# pyautogui.moveTo(100, 85)
# pyautogui.click(100, 85)
# pyautogui.click(100, 85)
# pyautogui.press('del')
# time.sleep(0.5)
# pyautogui.typewrite('1738856')
# pyautogui.press('enter')
# time.sleep(4)
#
# # Click to 'HASTA BAZINDA'
# pyautogui.moveTo(40, 300)
# pyautogui.click(40, 300)
# time.sleep(2)
#
# # Click to 'Laboratuvar'
# pyautogui.moveTo(400, 185)
# pyautogui.click(400, 185)
# time.sleep(2)
#
# # Search for the test.
# pyautogui.moveTo(90, 275)
# pyautogui.click(90, 275)
# pyautogui.typewrite('MENEN')
# pyautogui.press('enter')
#
# folder_path = 'C:\\Users\\eyas4\\Desktop\\TestTest'
# screenshot = pyautogui.screenshot(region=(60, 310, 200, 400))
# screenshot.save('C:\\Users\\eyas4\\Desktop\\TestTest\\screenshot.jpg')

image1 = cv2.imread('C:\\Users\\eyas4\\Desktop\\TestTest\\standard.jpg')
image2 = cv2.imread('C:\\Users\\eyas4\\Desktop\\TestTest\\screenshot.jpg')

image1_hsv = cv2.cvtColor(image1, cv2.COLOR_BGR2HSV)
image2_hsv = cv2.cvtColor(image2, cv2.COLOR_BGR2HSV)

hist1 = cv2.calcHist([image1_hsv], [0, 1], None, [180, 256], [0, 180, 0, 256])
hist2 = cv2.calcHist([image2_hsv], [0, 1], None, [180, 256], [0, 180, 0, 256])

correlation = cv2.compareHist(hist1, hist2, cv2.HISTCMP_CORREL)
chi_square = cv2.compareHist(hist1, hist2, cv2.HISTCMP_CHISQR)
intersection = cv2.compareHist(hist1, hist2, cv2.HISTCMP_INTERSECT)
bhattacharyya = cv2.compareHist(hist1, hist2, cv2.HISTCMP_BHATTACHARYYA)

if correlation == 1.0:
    print("Correlation: ", correlation)

print("Image Comparison Results:")
print("Correlation: ", correlation)
print("Chi-Square: ", chi_square)
print("Intersection: ", intersection)
print("Bhattacharyya Distance: ", bhattacharyya)
