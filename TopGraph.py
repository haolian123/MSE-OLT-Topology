
from graphviz import Digraph
import pandas as pd 

class TopologyDiagram:
    def __init__(self) :
        self.__root='0'


    
    # 获取所有节点中最多子节点的叶节点
    # 定义一个函数，用于获取树中最大叶子节点数
    def __maxLeavesNums(self,tree):
        # 获取当前树的叶子节点数
        leaves_nums = len(tree.keys())
        # 遍历树的每个节点
        for key, value in tree.items():
            # 如果节点的值是一个字典，表示还有子树
            if isinstance(value, dict):
                # 递归调用函数，获取子树的最大叶子节点数
                sum_leaves_nums = self.__maxLeavesNums(value)
                # 如果子树的最大叶子节点数大于当前树的叶子节点数，则更新最大叶子节点数
                if sum_leaves_nums > leaves_nums:
                    leaves_nums = sum_leaves_nums
        # 返回最大叶子节点数
        return leaves_nums
    

    # 定义一个函数，用于绘制树状图
    def __showTreemap(self,tree, name):
        # 创建一个有向图对象
        graph = Digraph("G", filename=name, format='png', strict=False)
        # 获取树的根节点
        root_node = list(tree.keys())[0]
        
        # 在图中添加根节点
        graph.node("0", root_node,fontname="Microsoft YaHei")
        # 调用辅助函数，递归绘制子树
        self.__drawTreemapHelper(graph, tree, "0")
        # 计算叶子节点数，用于设置图的布局
        leafs = str(self.__maxLeavesNums(tree) // 10)
        # 设置图的布局方向为从上到下
        graph.attr(rankdir='LR', ranksep=leafs)
        # 显示图形
        graph.render( view=False)


    
    # 定义一个辅助函数，用于递归绘制子树
    def __drawTreemapHelper(self,graph, tree, inc):
        # 全局变量，用于记录当前节点的编号
        # 获取子树的根节点
        root_node = list(tree.keys())[0]
        # 获取子树的所有子节点
        tree_nodes = tree[root_node]
        # 遍历子节点
        for i in tree_nodes.keys():
            # 如果子节点的值是一个字典，表示还有子树
            if isinstance(tree[root_node][i], dict):
                # 更新当前节点的编号
                self.__root = str(int(self.__root) + 1)
                # 在图中添加子节点
                graph.node(self.__root, list(tree[root_node][i].keys())[0],fontname="Microsoft YaHei")
                # 在图中添加边连接父节点和子节点
                graph.edge(inc, self.__root, str(i),fontname="Microsoft YaHei")
                # 递归调用函数，绘制子树
                self.__drawTreemapHelper(graph, tree[root_node][i], self.__root)
            else:
                # 更新当前节点的编号
                self.__root = str(int(self.__root) + 1)
                # 在图中添加叶子节点
                graph.node(self.__root, tree[root_node][i],fontname="Microsoft YaHei")
                # 在图中添加边连接父节点和叶子节点
                graph.edge(inc, self.__root, str(i),fontname="Microsoft YaHei")

    #主接口
    def MSE_to_OLT(self,file_path,sheet_name=2,MSE_name='Z端设备中文名称',OLT_name='A端设备中文名称'):

        #读取文件
        data=pd.read_excel(file_path,sheet_name=sheet_name)

        #提取需要的两列
        data=data[[MSE_name,OLT_name]]
       
        # data.sort_values(by=MSE_name)

        #去除重复行
        data=data.drop_duplicates(subset=[OLT_name])
        #去空值
        data=data.dropna(subset=[MSE_name])

        #按MSE分组
        dataGroupByMES=data.groupby(MSE_name)

        #存储所有MSE下所有的OLT的字典
        Tree={}

        #获得所有MSE的名称
        key_names=data[MSE_name].unique()

        #遍历所有MSE
        for mse_name in key_names:
            
            olt_names=list(dataGroupByMES.get_group(mse_name)[OLT_name])
            olt_cnt=1

            #存放不同前缀名下的所有OLT
            OLT_dict={}

            #得到同一前缀名的OLT
            for olt_name in olt_names:
                #提取OLT前缀
                olt_name_key=self.__get_olt_name(olt_name)
                #若字典还没有该MSE关键字
                if olt_name_key not in OLT_dict.keys():
                    OLT_dict[olt_name_key]=[]
                #将OLT加入列表
                OLT_dict[olt_name_key].append(olt_name)

            # print(OLT_dict.items())

            #拼装同一前缀名的OLT
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
            self.__showTreemap(tree_item,f'MSE-OLT/{key}/{key}.gv')

    #得到olt的前缀中文名
    def __get_olt_name(self,olt_name):
        index=olt_name.find('O')
        if index !=-1:
            res_name=olt_name[:index]
        else:
            res_name=olt_name
        return res_name


    


    

