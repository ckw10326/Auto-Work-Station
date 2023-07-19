x = 0
judge = 1
#設定答案，若x通過條件檢查，則設定judge = 0，跳出迴圈，進行下一項測試
while judge :
    #設定x為答案
    x = input("請輸入數字做為答案:")
    #由於input預設輸入值為字串，所以需要轉換為數字
    xx = int(x)
    if xx <10 and xx > 0:
        judge = 0

#設定y為遊戲玩家輸入數字，使用input功能
y = 0
count = 0 
judge = 1
while judge or count< 5:
    count = count + 1
    y = input("請輸入您猜的數字：")
    yy = int(y)
    #判斷玩家輸入數字是否符合答案
    if yy == xx:
        judge = 0
    if count == 5 :
        print("您已輸入錯誤超過5次，遊戲即將結束")
        judge = 0

print("遊戲結束")




