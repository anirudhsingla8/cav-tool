import pandas as pd
import openpyxl
import xlrd
import os
# common area
current_path=os.getcwd()
cav_excel='\\cav.xlsx'
other_tech_excel='\\other_tech.xlsx'
cable_assignment_FAB_excel='\\cable_assignment_FAB.xlsx'
cable_assignment_FAB_csv='\\cable_assignment_FAB.csv'
cable_assignment_excel='\\cable_assignment.xlsx'
cable_assignment_csv='\\cable_assignment.csv'
zero1_excel='\\zero1.xlsx'
cav_csv='\\cav.csv'
cable_csv="\\cable_FAB.csv"
cable_excel="\\cable_FAB.xlsx"
cable_with_gtd_excel='\\cable_with_gtd.xlsx'
cable_with_node_excel='\\cable_with_node.xlsx'
equipment_hfc_csv='\\equipment_hfc_FAB.csv'
equipment_node_excel='\\equipment_node.xlsx'
equipment_node_csv='\\equipment_node.csv'
common_node_csv='\\common_node.csv'
no_trench_excel='\\no_trench.xlsx'
final_excel='\\final.xlsx'
#equipment_gtd part
df = pd.read_csv(r'{0}{1}'.format(current_path,cable_csv))
cable_with_gtd=df[df['end_equipment_id'].str.contains("GTD",na=False)]
cable_with_gtd["id"].to_excel(r'{0}{1}'.format(current_path,cable_with_gtd_excel),index=False,header=1)
cable_gtd=pd.read_excel(r'{0}{1}'.format(current_path,cable_with_gtd_excel))
row_gtd=cable_gtd['id'].count()
comments=["cable ending on GTD"]*row_gtd
cable_gtd['comments']=comments
cable_gtd.to_excel(r'{0}{1}'.format(current_path,cable_with_gtd_excel),index=False,header=1)

#equipment_node part
df = pd.read_csv(r'{0}{1}'.format(current_path,equipment_hfc_csv))
equipment_node = df[df["hierarchy"] == "OTN"]
equipment_node["id"].to_excel(r'{0}{1}'.format(current_path,equipment_node_excel),index=False,header=1)

#equipment node part
cable = pd.read_excel(r'{0}{1}'.format(current_path,cable_excel))
cable['start_equipment_id']=cable.start_equipment_id.astype(str)
equipment_node=pd.read_excel('{0}{1}'.format(current_path,equipment_node_excel))
equipment_node['id']=equipment_node.id.astype(str)
cable_with_node=pd.merge(cable,equipment_node,left_on='start_equipment_id',right_on='id')
cable_with_node.rename(columns = {'id_x':'id'}, inplace = True)
cable_with_node["id"].to_excel('{0}{1}'.format(current_path,cable_with_node_excel),index=False,header=1)
cable_node=pd.read_excel('{0}{1}'.format(current_path,cable_with_node_excel))
row_count=cable_node['id'].count()
comments=["cable starting from node"]*row_count
cable_node['comments']=comments
cable_node.to_excel('{0}{1}'.format(current_path,cable_with_node_excel),index=False,header=1)

# zero length
zero1=pd.read_excel('{0}{1}'.format(current_path,cav_excel))
zero1=zero1[zero1['Asbuilt Length']==0]
zero1.rename(columns = {'Cable':'id'}, inplace = True)
zero1["id"].to_excel('{0}{1}'.format(current_path,zero1_excel),index=False,header=1)
zero2=pd.read_excel('{0}{1}'.format(current_path,zero1_excel))
comments=["zero length cable"]*zero2['id'].count()
zero2['comments']=comments
zero2.to_excel('{0}{1}'.format(current_path,zero1_excel),index=False,header=1)

# no_trench in T18
no_trench=pd.read_excel('{0}{1}'.format(current_path,cav_excel))
cable_assignment=pd.read_excel('{0}{1}'.format(current_path,cable_assignment_FAB_excel))
no_trench=no_trench[no_trench['Asbuilt Length']!=0]
no_trench=no_trench[no_trench['Issue']=='No Trench/Strand for segment']
no_trench.rename(columns = {'Cable':'id'}, inplace = True)
no_trench["id"].to_excel('{0}{1}'.format(current_path,no_trench_excel),index=False,header=1)
no_trench1=pd.read_excel('{0}{1}'.format(current_path,no_trench_excel))
cable_assignment.rename(columns={'cable_id':'id'},inplace=True)
cable_assignment['id'].to_excel('{0}{1}'.format(current_path,cable_assignment_excel),index=False,header=1)
cable_assignment=pd.read_excel('{0}{1}'.format(current_path,cable_assignment_excel))
common = no_trench1.merge(cable_assignment,on=['id'])
no_trench1=no_trench1[(~no_trench1.id.isin(common.id))&(~no_trench1.id.isin(common.id))]
comments=["no trench in T-eighteen"]*no_trench1['id'].count()
no_trench1['comments']=comments
no_trench1.to_excel('{0}{1}'.format(current_path,no_trench_excel),index=False,header=1)

# other technology
other_tech1=pd.read_excel('{0}{1}'.format(current_path,cav_excel))
other_tech2=pd.read_excel('{0}{1}'.format(current_path,cable_excel))
other_tech2=other_tech2[other_tech2['hierarchy']!='HDL']
other_tech3=pd.merge(other_tech1,other_tech2,left_on='Cable',right_on='id')
other_tech3.rename(columns = {'Cable':'id'}, inplace = True)
other_tech3['id'].to_excel('{0}{1}'.format(current_path,other_tech_excel),index=False,header=1)
other_tech3=pd.read_excel('{0}{1}'.format(current_path,other_tech_excel))
comments=['other technology']*other_tech3['id'].count()
other_tech3['comments']=comments
other_tech3.drop(columns=['id.1'],inplace=True)
other_tech3.to_excel('{0}{1}'.format(current_path,other_tech_excel),index=False,header=1)

#final merging
result1=pd.read_excel(r'{0}{1}'.format(current_path,cable_with_gtd_excel))
result2=pd.read_excel(r'{0}{1}'.format(current_path,cable_with_node_excel))
result3=pd.read_excel(r'{0}{1}'.format(current_path,zero1_excel))
result5=pd.read_excel(r'{0}{1}'.format(current_path,other_tech_excel))
result4=pd.read_excel(r'{0}{1}'.format(current_path,no_trench_excel))
frames=[result1,result2,result3,result4,result5]
result=pd.concat(frames)
result.to_excel('{0}{1}'.format(current_path,final_excel),index=False,header=1)

# final merging 2
final1=pd.read_excel('{0}{1}'.format(current_path,final_excel))
final2=pd.read_excel('{0}{1}'.format(current_path,cav_excel))
final2.rename(columns={'Cable':'id'}, inplace = True)
final=pd.merge(final1,final2,on='id',how='right')
final.rename(columns={'comments_x':'final_comments'}, inplace = True)
final.to_excel('{0}{1}'.format(current_path,final_excel),index=False,header=1)


os.remove(r'{0}{1}'.format(current_path,cable_with_node_excel))
os.remove(r'{0}{1}'.format(current_path,cable_with_gtd_excel))
os.remove(r'{0}{1}'.format(current_path,zero1_excel))
os.remove(r'{0}{1}'.format(current_path,other_tech_excel))
os.remove(r'{0}{1}'.format(current_path,no_trench_excel))
os.remove(r'{0}{1}'.format(current_path,equipment_node_excel))