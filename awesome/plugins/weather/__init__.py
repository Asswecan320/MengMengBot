from nonebot import on_command, CommandSession
import xlrd
import xlwt
import copy
from xlutils.copy import copy
import datetime
import style_config


#支出的分类
category = ['餐饮','家用','交通','娱乐','购物','人情','医疗','学习办公','子女财税','固定支出']



@on_command('支出', aliases=('记账','花销'))
async def turnover(session: CommandSession):

    macount = style_config.Acount()
    await session.aget(prompt='是晗还是锦QAQ？')
    who = session.current_arg_text
    while(1):
        if(who!='晗' and who!='锦' and who!='我们'):
            await session.aget(prompt='笨蛋输错啦，请重新输入!\n退出请输入0')
            who = session.current_arg_text
            if(who== 0):
                return
        else:
            if(who == '晗'):
                macount.Acount_list.append('晗')
                macount.user.font.colour_index = 71
            else:
                macount.Acount_list.append('锦')
                macount.user.font.colour_index = 17
            break
    macount.Acount_list.append(str(datetime.date.today()))
    await session.aget(prompt='花钱干嘛了捏要有分类的哦!\n'
                              '1.餐饮   2.家用   3.交通   4.娱乐\n'
                              '5.购物   6.人情   7.医疗  8.学习办公\n'
                              '9.子女财税   10.固定支出\n'
                              '退出请输入0哦')
    index = session.current_arg_text
    while (1):
        if(index.isdigit() and int(index)>=0 and int(index)<=10):
            if(index == 0):
                return
            else:
                macount.Acount_list += [category[int(index)-1]]
                break
        else:
            await session.aget(prompt='怎么还有笨蛋连数字都输不对啊？！给我重新输入！！(ps：还是0退出哦')
            index = session.current_arg_text
    await session.aget(prompt='好！现在摸着你的良心，你认为这项支出该划分到哪一个细类里呢？')
    macount.Acount_list += [session.current_arg_text]
    await session.aget(prompt='请输入金额____捏？')
    money = session.current_arg_text
    while(1):
        if (money.isdigit()):
            macount.Acount_list += [int(money)]
            break
        else:
            await session.aget(prompt='怎么还有笨蛋连数字都输不对啊？！给我重新输入！！(ps：还是0退出哦')
            money = session.current_arg_text
    await session.aget(prompt='努力点你马上就要完成这份记账了！现在来填最后一项吧~备注这笔支出哦，可以填无~')
    macount.Acount_list +=session.current_arg_text
    write_excel_streamTable(macount.Acount_list,'acount.xls',macount.user)
    await session.aget(prompt='本次记账完成咯~')
    return

def write_excel_streamTable(acount,file,user):
    ws=xlrd.open_workbook(filename = file,formatting_info = False)
    new_book = copy(ws)
    sheet = new_book.get_sheet(0)
    nrows = ws.sheet_by_index(0).nrows
    for i in range (len(acount)):
        sheet.write(nrows,i,acount[i],user)
    new_book.save('acount.xls')
    return
