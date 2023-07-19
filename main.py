from graphviz import Digraph
import pandas as pd 

# 获取所有节点中最多子节点的叶节点
# 定义一个函数，用于获取树中最大叶子节点数
def getMaxLeafs(myTree):
    # 获取当前树的叶子节点数
    numLeaf = len(myTree.keys())
    # 遍历树的每个节点
    for key, value in myTree.items():
        # 如果节点的值是一个字典，表示还有子树
        if isinstance(value, dict):
            # 递归调用函数，获取子树的最大叶子节点数
            sum_numLeaf = getMaxLeafs(value)
            # 如果子树的最大叶子节点数大于当前树的叶子节点数，则更新最大叶子节点数
            if sum_numLeaf > numLeaf:
                numLeaf = sum_numLeaf
    # 返回最大叶子节点数
    return numLeaf


# 定义一个函数，用于绘制树状图
def plot_model(tree, name):
    # 创建一个有向图对象
    g = Digraph("G", filename=name, format='png', strict=False)
    # 获取树的根节点
    first_label = list(tree.keys())[0]
    # print(first_label)
    # 在图中添加根节点
    g.node("0", first_label,fontname="Microsoft YaHei")
    # 调用辅助函数，递归绘制子树
    _sub_plot(g, tree, "0")
    # 计算叶子节点数，用于设置图的布局
    leafs = str(getMaxLeafs(tree) // 10)
    # 设置图的布局方向为从上到下
    g.attr(rankdir='LR', ranksep=leafs)
    # 显示图形
    g.render( view=False)

root="0"
# 定义一个辅助函数，用于递归绘制子树
def _sub_plot(g, tree, inc):
    # 全局变量，用于记录当前节点的编号
    
    global root
    # 获取子树的根节点
    first_label = list(tree.keys())[0]
    # 获取子树的所有子节点
    ts = tree[first_label]
    # 遍历子节点
    for i in ts.keys():
        # 如果子节点的值是一个字典，表示还有子树
        if isinstance(tree[first_label][i], dict):
            # 更新当前节点的编号
            root = str(int(root) + 1)
            # 在图中添加子节点
            g.node(root, list(tree[first_label][i].keys())[0],fontname="Microsoft YaHei")
            # 在图中添加边连接父节点和子节点
            g.edge(inc, root, str(i),fontname="Microsoft YaHei")
            # 递归调用函数，绘制子树
            _sub_plot(g, tree[first_label][i], root)
        else:
            # 更新当前节点的编号
            root = str(int(root) + 1)
            # 在图中添加叶子节点
            g.node(root, tree[first_label][i],fontname="Microsoft YaHei")
            # 在图中添加边连接父节点和叶子节点
            g.edge(inc, root, str(i),fontname="Microsoft YaHei")





data=pd.read_excel("揭阳OLT设备安全评估2023-05-29.xlsx",sheet_name=2)
data=data[['Z端设备中文名称','A端设备中文名称']]
data.sort_values(by='Z端设备中文名称')
data=data.drop_duplicates(subset=['A端设备中文名称'])
data=data.dropna(subset=['Z端设备中文名称'])


D=data.groupby('Z端设备中文名称')

#得到olt的前缀中文名
def get_olt_name(olt_name):
    index=olt_name.find('O')
    if index !=-1:
        res_name=olt_name[:index]
    else:
        res_name=olt_name
    return res_name

#得到所有MSE下所有的OLT
Tree={}
key_names=data['Z端设备中文名称'].unique()
key_names

for mse_name in key_names:
    olt_names=list(D.get_group(mse_name)['A端设备中文名称'])
    olt_cnt=1
    OLT_dict={}
    for olt_name in olt_names:
        #提取OLT前缀
        olt_name_key=get_olt_name(olt_name)

        if olt_name_key not in OLT_dict.keys():
            OLT_dict[olt_name_key]=[]
        OLT_dict[olt_name_key].append(olt_name)
    # print(OLT_dict.items())

    for key,values in OLT_dict.items():
        value=''
        for i in range(len(values)):
            value=value+values[i]
            if(i!=len(values)-1):
                value+=', '
            if (i+1)%3==0 and i!=0:
                value=value+'\n'
        if mse_name not in Tree.keys():
            Tree[mse_name]=dict()
        Tree[mse_name][str(olt_cnt)]=value
        olt_cnt+=1
    


    

for key,values in Tree.items():
    tree_item={key:values}
    # print(tree_item)
    plot_model(tree_item,f'MSE-OLT/{key}/{key}.gv')


