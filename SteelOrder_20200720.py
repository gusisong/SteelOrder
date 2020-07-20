import win32com.client, win32gui, win32con, subprocess
import time, datetime
import pandas as pd


def CreateOrder(account_name, password, data, apply_month):
    sap_app = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    subprocess.Popen(sap_app)
    time.sleep(2)

    flt = 0
    while flt == 0:
        try:
            hwnd = win32gui.FindWindow(None, "SAP Logon 720")
            flt = win32gui.FindWindowEx(hwnd, None, "Edit", None)  # capture handle of filter
        except:
            time.sleep(0.5)

    win32gui.SendMessage(flt, win32con.WM_SETTEXT, None, "svwsvp1a")
    win32gui.SendMessage(flt, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
    win32gui.SendMessage(flt, win32con.WM_KEYUP, win32con.VK_RIGHT, 0)
    time.sleep(0.1)

    dlg = win32gui.FindWindowEx(win32gui.FindWindow(None, "SAP Logon 720"), None, "Button", None)  # 登陆（0）
    win32gui.SendMessage(dlg, win32con.WM_LBUTTONDOWN, 0)
    win32gui.SendMessage(dlg, win32con.WM_LBUTTONUP, 0)

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = account_name
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
    session.findById("wnd[0]").sendVKey(0)

    # 进入创建采购申请界面
    session.findById("wnd[0]/tbar[0]/okcd").text = "ME51N"
    session.findById("wnd[0]").sendVKey(0)

    frames = []

    plant_list = list(set(data['工厂']))
    for plant in plant_list:

        # 设置缺省值（工厂、月份）
        try:
            session.findById(
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressToolbarButton(
                "&MEITPRP")
        except:
            session.findById(
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressToolbarButton(
                "&MEITPRP")

        session.findById("wnd[1]/usr/subSUB1:SAPLMEPERS:1106/ctxtMEREQ_PROP-WERKS").text = plant
        session.findById("wnd[1]/usr/subSUB1:SAPLMEPERS:1106/ctxtMEREQ_PROP-EEIND").text = ("{}.2020").format(
            apply_month)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # 筛选工厂
        fil_plant = data[data.工厂 == plant]

        # 筛选订货单位
        buyer_list = list(set(fil_plant['订货单位']))

        for buyer in buyer_list:
            fil_buyer = fil_plant[fil_plant.订货单位 == buyer]

            # 重置索引列
            fil_buyer = fil_buyer.reset_index(drop=True)

            # 自动生成分表(文件名不可使用特殊字符)
            # fil_buyer.to_csv('Detail\{0}_{1}_{2}月待申请数量.csv'.format(plant, buyer, apply_month), encoding='gb2312')

            # 填入物料号
            part_number_list = list(fil_buyer['物料编号'])
            count = 0
            for i in range(0, len(part_number_list), 10):
                part_number_group = part_number_list[i:i + 10]

                n = 0

                for part_number in part_number_group:
                    try:
                        session.findById(
                            "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                            count + n, "MATNR", part_number)

                    except:
                        session.findById(
                            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                            count + n, "MATNR", part_number)

                    n += 1

                try:
                    session.findById(
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter()
                except:
                    session.findById(
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter()

                count += 10

            # 校验基础信息是否完整，记录问题零件索引号
            failed_index_list = []
            for i in range(0, len(part_number_list)):
                try:
                    description = session.findById(
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").GetCellValue(
                        i, "TXZ01")
                except:
                    description = session.findById(
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").GetCellValue(
                        i, "TXZ01")

                if description:
                    continue
                else:
                    failed_index_list.append(i)

            # 如无问题零件，继续运行
            if failed_index_list == []:
                pass

            # 有问题零件，重新执行
            else:
                # 记录本轮问题零件Dataframe，合并入问题零件清单总表
                failed_df_new = fil_buyer.iloc[failed_index_list]
                frames.append(failed_df_new)

                # 更新fil_buyer，剔除问题零件
                fil_buyer = fil_buyer.drop(failed_index_list)

                # 退出并重新进入订单界面
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                session.findById("wnd[0]/tbar[0]/okcd").text = "ME51N"
                session.findById("wnd[0]").sendVKey(0)

                if fil_buyer.empty == True:
                    pass
                else:
                    # 填入物料号
                    part_number_list = list(fil_buyer['物料编号'])
                    count = 0
                    for i in range(0, len(part_number_list), 10):
                        part_number_group = part_number_list[i:i + 10]

                        n = 0

                        for part_number in part_number_group:
                            try:
                                session.findById(
                                    "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                                    count + n, "MATNR", part_number)

                            except:
                                session.findById(
                                    "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                                    count + n, "MATNR", part_number)

                            n += 1

                        try:
                            session.findById(
                                "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter()
                        except:
                            session.findById(
                                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter()

                        count += 10

            # 判断表格是否为空
            if fil_buyer.empty == True:
                pass

            else:
                # 填入订货数量
                order_volume_list = list(fil_buyer['订货'])
                count = 0
                for order_volume in order_volume_list:
                    try:
                        session.findById(
                            "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                            count, "MENGE", order_volume * 1000)
                    except:
                        session.findById(
                            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                            count, "MENGE", order_volume * 1000)
                    count += 1

                # 填入采购组织
                for i in range(0, len(order_volume_list)):
                    try:
                        session.findById(
                            "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                            i, "EKORG", "1000")
                    except:
                        session.findById(
                            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                            i, "EKORG", "1000")

                # 修正空缺价格为6
                for i in range(0, len(order_volume_list)):
                    try:
                        value = session.findById(
                            "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").GetCellValue(
                            i, "PREIS")
                    except:
                        value = session.findById(
                            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").GetCellValue(
                            i, "PREIS")

                    if float(value) == 0.00:
                        try:
                            session.findById(
                                "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                                i, "PREIS", "6")
                        except:
                            session.findById(
                                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell(
                                i, "PREIS", "6")
                    else:
                        continue

                # 回车确认
                try:
                    session.findById(
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter()
                except:
                    session.findById(
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter()

                # 正式启动用，保存提交采购订单
                # session.findById("wnd[0]/tbar[0]/btn[11]").press()

                localtime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                print('[{0}] {1}工厂_{2} 申请完成'.format(localtime, plant, buyer))

                # 调试用，不保存退出重进
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                session.findById("wnd[0]/tbar[0]/okcd").text = "ME51N"
                session.findById("wnd[0]").sendVKey(0)

                # 调试用，单次运行，循环运行需关闭
                # break

    # 问题零件清单写入文件
    if frames == []:
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        print('运行完毕，无问题零件')

    else:
        failed_df = pd.concat(frames)
        failed_df.to_csv(
            'FailedList/FailedList_{}.csv'.format(format(time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time())))),
            encoding='GB2312',
            index=None)

        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        print('运行完毕，{}个零件缺失基础信息未录入，清单见FailedList文件夹'.format(failed_df.shape[0]))


def main():
    # 读取SAP账号密码
    SapAccount = pd.read_csv('SAPaccount.csv')
    account_name = SapAccount.iloc[0, 0]
    password = SapAccount.iloc[0, 1]

    # 导入钢材订货计划数据
    data = pd.read_excel('OrderInfo/Template.xlsx', index_col=0, header=1)
    data.rename(columns={'料片\n供应商': '订货单位'}, inplace=True)

    # 设定采购订单申请月份
    apply_month = input('请输入订单申请月份(数字)，按回车')
    # apply_month = datetime.datetime.now().month

    CreateOrder(account_name, password, data, apply_month)


if __name__ == '__main__':
    main()
