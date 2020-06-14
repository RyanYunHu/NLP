# coding:utf-8
import pandas as pd
import jieba
import jieba.posseg as pseg

path_c = r'E:\网仓科技\促销特征值计算\资料文档\mathch\out.csv'
path_p = r'E:\网仓科技\促销特征值计算\促销数据自动化提取\促销因素及门店信息.xlsx'
data = pd.read_csv(path_c, encoding='gb18030', engine='python')
data_promotion = pd.read_excel(path_p)

data_cate = set(data['品类名称'])
data_bran = set(data['品牌名称'])
data_pric = set(data['价格分类名称'])
data_sus = []
for i in range(len(data['系列名称1'])):
    if data['系列名称1'][i]:
        for j in range(str(data['系列名称1'][i]).count('&')):
            data_sus.append(data['系列名称1'][i].split('&')[j + 1])
data_subs = set(data_sus)

data_suc = []
for i in range(len(data['子品类名称1'])):
    if data['子品类名称1'][i]:
        for j in range(str(data['子品类名称1'][i]).count('&')):
            data_suc.append(data['子品类名称1'][i].split('&')[j + 1])
data_subc = set(data_suc)

data_prom_flow = set(data_promotion['促销因素流量'])
data_prom_rero = data_promotion['促销因素转化率'].where(data_promotion['促销因素转化率'].notnull(), '&')
data_prom_rero1 = list(set(data_prom_rero))
data_prom_rero1.remove('&')

ds = pd.read_csv(r'C:\Users\yunrui.hu\Desktop\test\a.csv', encoding='gbk')

#加载自定义词库
file = open(r'C:\Users\yunrui.hu\Desktop\userdict.txt',encoding='UTF-8')
user_dict = file.read()
jieba.load_userdict(r'C:\Users\yunrui.hu\Desktop\userdict.txt')

ds_re = []
ds_su = []
for i in range(len(ds)):
    ds_re.append([(x.word, x.flag) for x in pseg.cut(ds['活动要求'][i])])
    ds_su.append([(x.word, x.flag) for x in pseg.cut(ds['活动支持'][i])])

ds_re_pos = []
ds_su_pos = []
for i in range(len(ds)):
    ds_re_pos.append([(x[0], x[1], x[2]) for x in jieba.tokenize(ds['活动要求'][i])])
    ds_su_pos.append([(x[0], x[1], x[2]) for x in jieba.tokenize(ds['活动支持'][i])])

ds['活动要求分词'] = ds_re
ds['活动支持分词'] = ds_su
ds['活动要求分词位置'] = ds_re_pos
ds['活动支持分词位置'] = ds_su_pos


ds_re_cate = []
ds_re_bran = []
ds_re_pric = []
ds_re_subs = []
ds_re_subc = []
ds_re_prom_flow = []
ds_re_prom_rero1 = []

ds_su_cate = []
ds_su_bran = []
ds_su_pric = []
ds_su_subs = []
ds_su_subc = []
ds_su_prom_flow = []
ds_su_prom_rero1 = []


ds['活动要求中的品类'] = None
ds['活动要求中的品牌'] = None
ds['活动要求中的大系列'] = None
ds['活动要求中的子系列'] = None
ds['活动要求中的子品类'] = None
ds['活动要求中的促销因素流量'] = None
ds['活动要求中的促销因素转化率'] = None

ds['活动支持中的品类'] = None
ds['活动支持中的品牌'] = None
ds['活动支持中的大系列'] = None
ds['活动支持中的子系列'] = None
ds['活动支持中的子品类'] = None
ds['活动支持中的促销因素流量'] = None
ds['活动支持中的促销因素转化率'] = None

for i in range(len(ds['活动要求'])):
    for j1 in range(len(data_cate)):
        if str(list(data_cate)[j1]) in ds['活动要求'][i]:
            ds_re_cate.append(list(data_cate)[j1])
        else:
            ds_re_cate.append(' ')
    ds['活动要求中的品类'][i] = ds_re_cate
    ds_re_cate = []

    for j2 in range(len(data_bran)):
        if str(list(data_bran)[j2]) in ds['活动要求'][i]:
            ds_re_bran.append(list(data_bran)[j2])
        else:
            ds_re_bran.append(' ')
    ds['活动要求中的品牌'][i] = ds_re_bran
    ds_re_bran = []

    for j3 in range(len(data_pric)):
        if str(list(data_pric)[j3]) in ds['活动要求'][i]:
            ds_re_pric.append(list(data_pric)[j3])
        else:
            ds_re_pric.append(' ')
    ds['活动要求中的大系列'][i] = ds_re_pric
    ds_re_pric = []

    for j4 in range(len(data_subs)):
        if str(list(data_subs)[j4]) in ds['活动要求'][i]:
            ds_re_subs.append(list(data_subs)[j4])
        else:
            ds_re_subs.append(' ')
    ds['活动要求中的子系列'][i] = ds_re_subs
    ds_re_subs = []

    for j5 in range(len(data_subc)):
        if str(list(data_subc)[j5]) in ds['活动要求'][i]:
            ds_re_subc.append(list(data_subc)[j5])
        else:
            ds_re_subc.append(' ')
    ds['活动要求中的子品类'][i] = ds_re_subc
    ds_re_subc = []

    for j6 in range(len(data_prom_flow)):
        if str(list(data_prom_flow)[j6]) in ds['活动要求'][i]:
            ds_re_prom_flow.append(list(data_prom_flow)[j6])
        else:
            ds_re_prom_flow.append(' ')
    ds['活动要求中的促销因素流量'][i] = ds_re_prom_flow
    ds_re_prom_flow = []

    for j7 in range(len(data_prom_rero1)):
        if str(list(data_prom_rero1)[j7]) in ds['活动要求'][i]:
            ds_re_prom_rero1.append(list(data_prom_rero1)[j7])
        else:
            ds_re_prom_rero1.append(' ')
    ds['活动要求中的促销因素转化率'][i] = ds_re_prom_rero1
    ds_re_prom_rero1 = []


    for k1 in range(len(data_cate)):
        if str(list(data_cate)[k1]) in ds['活动支持'][i]:
            ds_su_cate.append(list(data_cate)[k1])
        else:
            ds_su_cate.append(' ')
    ds['活动支持中的品类'][i] = ds_su_cate
    ds_su_cate = []

    for k2 in range(len(data_bran)):
        if str(list(data_bran)[k2]) in ds['活动支持'][i]:
            ds_su_bran.append(list(data_bran)[k2])
        else:
            ds_su_bran.append(' ')
    ds['活动支持中的品牌'][i] = ds_su_bran
    ds_su_bran = []

    for k3 in range(len(data_pric)):
        if str(list(data_pric)[k3]) in ds['活动支持'][i]:
            ds_su_pric.append(list(data_pric)[k3])
        else:
            ds_su_pric.append(' ')
    ds['活动支持中的大系列'][i] = ds_su_pric
    ds_su_pric = []

    for k4 in range(len(data_subs)):
        if str(list(data_subs)[k4]) in ds['活动支持'][i]:
            ds_su_subs.append(list(data_subs)[k4])
        else:
            ds_su_subs.append(' ')
    ds['活动支持中的子系列'][i] = ds_su_subs
    ds_su_subs = []

    for k5 in range(len(data_subc)):
        if str(list(data_subc)[k5]) in ds['活动支持'][i]:
            ds_su_subc.append(list(data_subc)[k5])
        else:
            ds_su_subc.append(' ')
    ds['活动支持中的子品类'][i] = ds_su_subc
    ds_su_subc = []

    for k6 in range(len(data_prom_flow)):
        if str(list(data_prom_flow)[k6]) in ds['活动支持'][i]:
            ds_su_prom_flow.append(list(data_prom_flow)[k6])
        else:
            ds_su_prom_flow.append(' ')
    ds['活动支持中的促销因素流量'][i] = ds_su_prom_flow
    ds_su_prom_flow = []

    for k7 in range(len(data_prom_rero1)):
        if str(list(data_prom_rero1)[k7]) in ds['活动支持'][i]:
            ds_su_prom_rero1.append(list(data_prom_rero1)[k7])
        else:
            ds_su_prom_rero1.append(' ')
    ds['活动支持中的促销因素转化率'][i] = ds_su_prom_rero1
    ds_su_prom_rero1 = []


ds['要求中的商品信息关键词'] = (ds['活动要求中的品类'] + ds['活动要求中的品牌'] + ds['活动要求中的大系列'] + ds['活动要求中的子系列'] + ds['活动要求中的子品类'])
ds['支持中的商品信息关键词'] = (ds['活动支持中的品类'] + ds['活动支持中的品牌'] + ds['活动支持中的大系列'] + ds['活动支持中的子系列'] + ds['活动支持中的子品类'])
ds['要求中的促销流量信息关键词'] = ds['活动要求中的促销因素流量']
ds['要求中的促销转化率信息关键词'] = ds['活动要求中的促销因素转化率']
ds['支持中的促销流量信息关键词'] = ds['活动支持中的促销因素流量']
ds['支持中的促销转化率信息关键词'] = ds['活动支持中的促销因素转化率']

for i in range(len(ds)):
    for j in range(ds['要求中的商品信息关键词'][i].count(' ')):
        ds['要求中的商品信息关键词'][i].remove(' ')
    for j in range(ds['活动要求中的品类'][i].count(' ')):
        ds['活动要求中的品类'][i].remove(' ')
    for j in range(ds['活动要求中的品牌'][i].count(' ')):
        ds['活动要求中的品牌'][i].remove(' ')
    for j in range(ds['活动要求中的大系列'][i].count(' ')):
        ds['活动要求中的大系列'][i].remove(' ')
    for j in range(ds['活动要求中的子系列'][i].count(' ')):
        ds['活动要求中的子系列'][i].remove(' ')
    for j in range(ds['活动要求中的子品类'][i].count(' ')):
        ds['活动要求中的子品类'][i].remove(' ')
    for j in range(ds['支持中的商品信息关键词'][i].count(' ')):
        ds['支持中的商品信息关键词'][i].remove(' ')
    for j in range(ds['活动支持中的品类'][i].count(' ')):
        ds['活动支持中的品类'][i].remove(' ')
    for j in range(ds['活动支持中的品牌'][i].count(' ')):
        ds['活动支持中的品牌'][i].remove(' ')
    for j in range(ds['活动支持中的大系列'][i].count(' ')):
        ds['活动支持中的大系列'][i].remove(' ')
    for j in range(ds['活动支持中的子系列'][i].count(' ')):
        ds['活动支持中的子系列'][i].remove(' ')
    for j in range(ds['活动支持中的子品类'][i].count(' ')):
        ds['活动支持中的子品类'][i].remove(' ')

    for j in range(ds['要求中的促销流量信息关键词'][i].count(' ')):
        ds['要求中的促销流量信息关键词'][i].remove(' ')
    for j in range(ds['要求中的促销转化率信息关键词'][i].count(' ')):
        ds['要求中的促销转化率信息关键词'][i].remove(' ')
    for j in range(ds['支持中的促销流量信息关键词'][i].count(' ')):
        ds['支持中的促销流量信息关键词'][i].remove(' ')
    for j in range(ds['支持中的促销转化率信息关键词'][i].count(' ')):
        ds['支持中的促销转化率信息关键词'][i].remove(' ')

no_re = 0
num_re_word = 0
no_su = 0
num_su_word = 0
for i in range(len(ds)):
    num_re_word = num_re_word + len(ds['活动要求分词'][i])
    num_su_word = num_su_word + len(ds['活动支持分词'][i])

df_re = pd.DataFrame([], columns=['行号', '词', '词性', '起始位置', '末尾位置', '成分'], index=range(num_re_word))
df_su = pd.DataFrame([], columns=['行号', '词', '词性', '起始位置', '末尾位置', '成分'], index=range(num_su_word))

for i in range(len(ds)):
    for j in range(len(ds['活动要求分词'][i])):
        if ds['活动要求分词'][i][j][1].strip(' ') != 'x':
            df_re['行号'][no_re] = str(i)
            df_re['词'][no_re] = ds['活动要求分词'][i][j][0]
            df_re['词性'][no_re] = ds['活动要求分词'][i][j][1]
            df_re['起始位置'][no_re] = (ds['活动要求'][i]).find(ds['活动要求分词'][i][j][0])
            df_re['末尾位置'][no_re] = (ds['活动要求'][i]).find(ds['活动要求分词'][i][j][0]) + len(ds['活动要求分词'][i][j][0])
            if ds['活动要求分词'][i][j][0] in ds['活动要求中的品类'][i]:
                df_re['成分'][no_re] = '品类'
            elif ds['活动要求分词'][i][j][0] in ds['活动要求中的品牌'][i]:
                df_re['成分'][no_re] = '品牌'
            elif ds['活动要求分词'][i][j][0] in ds['活动要求中的大系列'][i]:
                df_re['成分'][no_re] = '大系列'
            elif ds['活动要求分词'][i][j][0] in ds['活动要求中的子系列'][i]:
                df_re['成分'][no_re] = '子系列'
            elif ds['活动要求分词'][i][j][0] in ds['活动要求中的子品类'][i]:
                df_re['成分'][no_re] = '子品类'
            elif ds['活动要求分词'][i][j][0] in ds['活动要求中的促销因素流量'][i]:
                df_re['成分'][no_re] = '促销因素流量'
            elif ds['活动要求分词'][i][j][0] in ds['活动要求中的促销因素转化率'][i]:
                df_re['成分'][no_re] = '促销因素转化率'
            elif ds['活动要求分词'][i][j][1].strip(' ') == 'm' and len(str(ds['活动要求分词'][i][j][0])) >= 8:
                df_re['成分'][no_re] = '商品编码or条码'
            else:
                df_re['成分'][no_re] = ''
            no_re = no_re + 1

for i in range(len(ds)):
    for j in range(len(ds['活动支持分词'][i])):
        if ds['活动支持分词'][i][j][1].strip(' ') != 'x':
            df_su['行号'][no_su] = str(i)
            df_su['词'][no_su] = ds['活动支持分词'][i][j][0]
            df_su['词性'][no_su] = ds['活动支持分词'][i][j][1]
            df_su['起始位置'][no_su] = (ds['活动支持'][i]).find(ds['活动支持分词'][i][j][0])
            df_su['末尾位置'][no_su] = (ds['活动支持'][i]).find(ds['活动支持分词'][i][j][0]) + len(ds['活动支持分词'][i][j][0])
            if ds['活动支持分词'][i][j][0] in ds['活动支持中的品类'][i]:
                df_su['成分'][no_su] = '品类'
            elif ds['活动支持分词'][i][j][0] in ds['活动支持中的品牌'][i]:
                df_su['成分'][no_su] = '品牌'
            elif ds['活动支持分词'][i][j][0] in ds['活动支持中的大系列'][i]:
                df_su['成分'][no_su] = '大系列'
            elif ds['活动支持分词'][i][j][0] in ds['活动支持中的子系列'][i]:
                df_su['成分'][no_su] = '子系列'
            elif ds['活动支持分词'][i][j][0] in ds['活动支持中的子品类'][i]:
                df_su['成分'][no_su] = '子品类'
            elif ds['活动支持分词'][i][j][0] in ds['活动支持中的促销因素流量'][i]:
                df_su['成分'][no_su] = '促销因素流量'
            elif ds['活动支持分词'][i][j][0] in ds['活动支持中的促销因素转化率'][i]:
                df_su['成分'][no_su] = '促销因素转化率'
            elif ds['活动支持分词'][i][j][1].strip(' ') == 'm' and len(ds['活动支持分词'][i][j][0]) >= 8:
                df_su['成分'][no_su] = '商品编码or条码'
            else:
                df_su['成分'][no_su] = ''
            no_su = no_su + 1



ds.drop('Unnamed: 0', 1).to_csv(r'C:\Users\yunrui.hu\Desktop\temp.csv', encoding='gbk', index=False)
df_re.to_csv(r'C:\Users\yunrui.hu\Desktop\temp1.csv', encoding='gbk', index=False)
df_su.to_csv(r'C:\Users\yunrui.hu\Desktop\temp2.csv', encoding='gbk', index=False)


