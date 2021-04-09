create table #tmpCRMImport(ID nvarchar(255),
CusCode nvarchar(255))

insert #tmpCRMImport values('e6aa9d03-4408-41fe-9628-81fc7390a89e_CKH0912040','CKH0912040')
insert #tmpCRMImport values('a887d99a-694c-444e-a85e-9e8ea16d2afc_CKH0834018','CKH0834018')
insert #tmpCRMImport values('464bea57-4e80-47be-82f2-d8da15742aa4_CKH0812023','CKH0812023')
insert #tmpCRMImport values('1e252aeb-a802-49e1-b642-bfb72e6559b9_CKH0512022','CKH0512022')
insert #tmpCRMImport values('aff4fdb6-9b12-4b37-8ce1-cc7e3354c6e4_CKH0028031','CKH0028031')
insert #tmpCRMImport values('356ab420-80f1-4793-b8c9-8b75dce8f53f_CKH0411005','CKH0411005')

select * from #tmpCRMImport
go



select * from tbiz_autharea_c_2012 

--delete from tbiz_autharea_c_2012
--where Id in (select ID from #tmpCRMImport)

select * from tbiz_saletarget_c_2012 

--delete from tbiz_saletarget_c_2012 
--where Id in (select ID from #tmpCRMImport)

select * from [dbo].[tbiz_autharea_2012]

--delete from [dbo].[tbiz_autharea_2012]
--where Id in (select ID from #tmpCRMImport)

select * from [dbo].[tbiz_saletarget_2012]

--delete from [dbo].[tbiz_saletarget_2012]
--where Id in (select ID from #tmpCRMImport)

select * from [TPRT_Client_TYPE]
--delete from [dbo].[TPRT_Client_TYPE]
--where Last_UPDATE = '2021/03/01'

