def Text_Generator(def_l_title, def_l_vision, def_l_Num, Letter_Date, PlanNo):
    Plan_No = int(PlanNo)
    judge00 = input("是否有意見，有請輸入1，無請輸入0：")
    if int(judge00):
        My_company_Num = str(input("請輸入審查意見填寫單位:\n 1 = 南部施工處 \n 2 = 中部施工處 \n 3 = 興達發電廠 \n 4 =台中發電廠"))
        Pages = str(input("請輸入審查意見頁數:"))
        My_company = ["", "南部施工處", "中部施工處", "興達發電廠", "台中發電廠"]
        Consult_Company = ["", "吉興公司", "泰興公司", "GE/CTCI"]
        
        #複製套印文件
        Print_page_File = dest_folder + r"\套印_" + def_l_Num + ".rtf"
        print(Pring_page_StandardFile)
        print(dest_folder)
        print(Print_page_File)
        input("plz enter any key")
        shutil.copy(Pring_page_StandardFile, Print_page_File)
        #複製傳真文件
        Fax_PageFile = dest_folder + r"\Fax_" + def_l_Num + ".doc"
        shutil.copy(Source_Fax_PageFile, Fax_PageFile)
        
        date_obj = datetime.datetime.strptime(Letter_Date, "%Y/%m/%d")
        month = date_obj.strftime("%m")
        day = date_obj.strftime("%d")
        Plan_Name = "興達" if "CTC" in def_l_Num else "台中"

        contents0 = "本文係" + My_company[My_company_Num] + "對統包商提送「" + def_l_title + "」" + "Rev." + str(def_l_vision) +"所提審查意見(共" + Pages + "頁)" + "，未逾合約規範，已電傳" + Consult_Company[Plan_No] + "，擬陳閱後文存。"
        contents1 = "檢送" + Plan_Name + "電廠燃氣機組更新改建計畫" + def_l_title + "Rev." + str(def_l_vision) + "，" + My_company[My_company_Num] + "之審查意見（如附，共" + Pages + "頁）供卓參，請查照。"
        contents2 = "依據GE/CTCI 112年" + month + "月" + day +"日" + def_l_Num + "號辦理。"
        contents4 = "本文係統包商提送「" + def_l_title + "」" + "Rev." + str(def_l_vision) +"，本組無意見，已Email通知" + Consult_Company[Plan_No] + "公司" + "，擬陳閱後文存。"
        print("----------------他單位審查意見簽辦------------------")
        print(contents0)
        print("----------------傳真------------------")
        print(contents1)
        print(contents2)
        print("----------------主辦簽辦------------------")
        print(contents4)
        input("暫停")

    return 0