


g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")

# 미체결 조회 서비스
obj = Cp5339()


# 미체결 주문 조회
def Reqeust5339():
    diOrderList = {}
    orderList = []
    obj.Request5339(diOrderList, orderList)
    print("[Reqeust5339]미체결 개수 ", len(orderList))


Reqeust5339()