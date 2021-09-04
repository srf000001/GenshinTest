import re

Judgeline = [r'^\S{1,4}$', r'^[1-9]\d{0,4}$', r'^[1-9]\d{0,3}$', r'^[1-9]\d{0,2}$', r'^[1-9]\d{0,2}$',
             r'^[1-9]\d{0,2}$', r'^[1-9]\d{0,2}$',
             r'^[1-9]\d{0,3}$', r'^[1-9]\d{0,3}$']

if __name__ == '__main__':

    print(len('魈'))
    m1 = re.match(Judgeline[1], '111111')
    if m1:
        print('配对成功')
    else:
        print('配队失败')
