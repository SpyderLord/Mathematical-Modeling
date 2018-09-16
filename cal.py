import final as final
import xlrd
import xlwt

def get_injury_value(value):
    '''

    :param value: 伤亡总数
    :return: 根据伤亡总数对应的区间数返回对应的结果
    '''
    if value==0:
        return 0
    elif value<2001:
        print("injury num=%f\n"%final.injury_total[0])
        return final.injury_total[0]
    elif value<4001:
        print("injury num=%f\n" % final.injury_total[1])
        return final.injury_total[1]
    elif value<6001:
        print("injury num=%f\n" % final.injury_total[2])
        return final.injury_total[2]
    elif value<8001:
        print("injury num=%f\n" % final.injury_total[3])
        return final.injury_total[3]
    else:
        print("injury num=%f\n" % final.injury_total[4])
        return final.injury_total[4]

def get_death_total(value):
    '''

    :param value:死亡人数
    :return: 根据死亡人数返回对应的区间值
    '''
    if value==0:
        return 0
    elif value<401:
        print("death num=%f\n" % final.death_total[0])
        return final.death_total[0]
    elif value<801:
        print("death num=%f\n" % final.death_total[1])
        return final.death_total[1]
    elif value<1201:
        print("death num=%f\n" % final.death_total[2])
        return final.death_total[2]
    elif value<1601:
        print("death num=%f\n" % final.death_total[3])
        return final.death_total[3]
    else:
        print("death num=%f\n" % final.death_total[4])
        return final.death_total[4]

def get_murderer_num(value):
    '''

    :param value:凶手数量
    :return:根据凶手数量得到对应的区间值
    '''
    if value==0:
        return 0
    elif value<1001:
        print("murderer num=%f\n" % final.murderer_num[0])
        return final.murderer_num[0]
    elif value<2001:
        print("muderer num=%f\n" % final.murderer_num[1])
        return final.murderer_num[1]
    elif value<3001:
        print("murderer num=%f\n" % final.murderer_num[2])
        return final.murderer_num[2]
    elif value<4001:
        print("murderer num=%f\n" % final.murderer_num[3])
        return final.murderer_num[3]
    else:
        print("murderer num=%f\n" % final.murderer_num[4])
        return final.murderer_num[4]

def get_hostage_value(value):
    '''

    :param value: 事件中是否存在人质
    :return: 根据事件中是否存在人质返回对应的区间值
    '''
    if value==1:
        print("hostage num=%f\n" % final.hostage[0])
        return final.hostage[0]
    elif value==0:
        print("hostage num=%f\n" % final.hostage[1])
        return final.hostage[1]
    else:
        print("hostage num=%f\n" % final.hostage[2])
        return final.hostage[2]

def get_area_code(index):
    '''

    :param value:地区代码
    :return: 根据地区代码返回对应的区间值
    '''
    print("area_code num=%f\n" % final.area_code[(int)(index-1)])
    return final.area_code[(int)(index-1)]

def get_caichan_damage(index):
    '''

    :param index:地区索引值
    :return: 根据地区索引值返回对应的区间值
    '''
    if index==0:
        return 0
    else:
        print("caichan_damage num=%f\n" % final.caichan_damage[(int)(index-1)])
        return final.caichan_damage[(int)(index-1)]

def get_attack_type(index):
    '''

    :param index:攻击类型
    :return: 根据攻击类型返回对应的区间值
    '''
    if index==9:
        print("attack_type num=0")
        return 0;
    else:
        print("attack_type num=%f\n" % final.attack_type[(int)(index-1)])
        return final.attack_type[(int)(index-1)]

def get_second_attack(index):
    '''

    :param index:
    :return:
    '''
    if index==0:
        print("second attack num=0")
        return 0
    else:
        print("second attack num=%f"%get_attack_type(index))
        return get_attack_type(index)


def get_third_attack(index):
    '''

    :param index:
    :return:
    '''
    if index == 0:
        print("third attack num=0")
        return 0
    else:
        print("third attack num=%f"%get_attack_type(index))
        return get_attack_type(index)

def get_target_victim(index):
    '''

    :param index:目标受害者类型
    :return:根据目标受害者类型返回对应的区间值
    '''
    print("target victim num=%f" % final.target_victim[(int)(index-1)])
    return final.target_victim[(int)(index-1)]

def get_weapon_type(index):
    '''

    :param index:武器类型
    :return: 根据武器类型返回对应的区间值
    '''
    print("weapon type num=%f"%final.weapon_type[(int)(index-1)])
    return final.weapon_type[(int)(index-1)]

def get_second_weapon(index):
    '''

    :param index:
    :return:
    '''
    if index==0:
        print("second weapon type=0")
        return 0
    else:
        print("second weapon type=%f"%get_weapon_type(index))
        return get_weapon_type(index)


def get_third_weapon(index):
    '''

    :param index:
    :return:
    '''
    if index == 0:
        print("third weapon type=0")
        return 0
    else:
        print("third weapon type=%f" % get_weapon_type(index))
        return get_weapon_type(index)

def calculate_final_value(row_value):
    '''

    :param row_value:excel中一行的数据
    :return: 最终的计算结果
    '''
    final_value=final.final[0]*row_value[0]\
                +final.final[1]*row_value[1]\
                +final.final[2]*get_area_code(row_value[2])\
                +final.final[3]*row_value[3]\
                +final.final[4]*row_value[4]\
                +final.final[5]*row_value[5]\
                +final.final[6]*row_value[6]\
                +final.final[7]*row_value[7]\
                +final.final[8]*get_attack_type(row_value[8])\
                +final.final[9]*get_second_attack(row_value[9])\
                +final.final[10]*get_third_attack(row_value[10])\
                +final.final[11]*get_target_victim(row_value[11])\
                +final.final[12]*row_value[12]\
                +final.final[13]*row_value[13]\
                +final.final[14]*get_murderer_num(row_value[14])\
                +final.final[15]*get_weapon_type(row_value[15])\
                +final.final[16]*get_second_weapon(row_value[16])\
                +final.final[17]*get_third_weapon(row_value[17])\
                +final.final[18]*get_death_total(row_value[18])\
                +final.final[19]*get_injury_value(row_value[19])\
                +final.final[20]*get_caichan_damage(row_value[20])\
                +final.final[21]*get_hostage_value(row_value[21])
    return final_value

if __name__=='__main__':
    workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = workbook.add_sheet('final', cell_overwrite_ok=True)
    sheet.write(0, 1, "value")

    #读取原始数据
    data = xlrd.open_workbook(r'./选择-6万以后.xlsx')
    table = data.sheets()[0]
    ncols = table.ncols
    nrows = table.nrows
    print("rows%d"%nrows)
    for i in range(nrows-1):
        print("index=%d"%(i+1))
        row_value=table.row_values(i+1)
        final_value=calculate_final_value(row_value)
        sheet.write(i+1,1,final_value)
    workbook.save(r'./final2.xlsx')