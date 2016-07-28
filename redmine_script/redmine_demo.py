from redmine import Redmine
from datetime import date

#创建redmine连接
# redmine=Redmine('http://demo.redmine.org/',raise_attr_exception=False)
redmine=Redmine('http://demo.redmine.org/',username='liuhq419',password='hq2070.com')
print(redmine)
#获得项目标识为liuhq419的项目
project=redmine.project.get("liuhq419")  #liuhq419是项目的标识，唯一性，并非项目的名称。
print('project_id:',project.id)
# print(project.identifier)
# print(project.created_on)
#print(project.issues)
#print(project.issues[0])
#dir(project.issues[0])
#dir(redmine.project.get('liuhq419'))

#创建项目 --失败
# project=redmine.project.create(name='Vacation',identifier='vacation',description='foo',
#                                homepage='http://liuhq419.bar',is_public=True,parent_id=345,inherit_members=True,
#                                custom_fields=[{'id':1,'value':'foo'},{'id':2,'value':'bar'}])
# print(project)
# project=redmine.project.new()
# project.name='liuhaiqin'
# project.identifier='liuhaiqin'
# project.description='liu'
# project.homepage='http://liuhq.liu'
# project.is_public=True
# project.parent_id=345
# project.inherit_members=True
# project.custom_fields=[{'id':1,'value':'foo'},{'id':2,'value':'bar'}]
# project.save()

'''
l=list(redmine.issue_status.all(limit=1))
print(l)
v=list(redmine.issue_status.all(limit=1).values())
print(v)
pros=redmine.project.all().values()
print(pros)
#print(pros.values())


#创建issue --create
#project_id 是项目的标识或者id
#project_id='1153539',或者project_id='liuhq419'
issue=redmine.issue.create(project_id='115359',subject='test',track_id=8,description='foo',status_id=3,priority_id=7,
                           assigned_to_id=123,start_date=date(2016,7,1),due_date=date(2016,8,1),estimated_hours=3,
                           done_ratio=40,custom_fields=[{'id': 1, 'value': 'foo'}, {'id': 2, 'value': 'bar'}],)
                          # uploads= [{'path': '/absolute/path/to/file'}, {'path': '/absolute/path/to/file2'}],)
print(issue)

##创建issue --new
issue=redmine.issue.new()
issue.project_id='liuhq419'
issue.subject='nihao'
issue.track_id=4
issue.description='nihaoma'
issue.statues_id=4
issue.priority=4
issue.assigned_to_id=3
issue.start_date=date(2016,1,1)
issue.due_date=date(2016,7,20)
issue.estimated_hour=10
issue.save()
print(issue)
print(issue.id)   id=141788'''

# issue=redmine.issue.get(141788)
# print(issue)
# print(issue.journals)
#
# issue1=redmine.issue.get(141788,include='children,journals,watchers')
# print(issue1)


# all
#包括三个方法：sort,limit,offset
# issues=redmine.issue.all(sort='catagory:desc',limit=2)
#print(issues)

#filter
#project_id,subject_id,track_id,status_id,query_id,assigned_id, cf_x,sort,limit.offset
# issues=redmine.issue.filter(project_id='liuhq419',subject_id='nihao',
#                             create_on='><2016-07-15|2016-07-18',sort='category:desc')
# print(issues)

#update  #没有权限
# b=redmine.issue.update(141788,project_id='liuhq419',subject='nihao')
# print(b)


#watcher
#add watcher
# issue=redmine.issue.get(141788)
# b=issue.watcher.add(1)
# print(b)

#project Membership
#create methods
# membership=redmine.project_membership.create(project_id='liuhq419',user_id=1,role_ids=[1,2])
# print(membership)

# read methods
#get
# m=project.memberships[0]
# print('m:',m)
# membership=redmine.project_membership.get(156303)
# print('membership:',membership)
#
# # filter
# memberships=redmine.project_membership.filter(project_id='liuliu')
# print(memberships)

#user  --get  没有权限
user1=redmine.user.get(1,include='membership,groups')
print('user:',user1)

current_user=redmine.user.get('current')
print('current_user',current_user)