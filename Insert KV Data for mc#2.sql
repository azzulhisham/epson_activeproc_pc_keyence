use marking

update Setting set spec = 'FA-128_000'
where ctrlno = 'M02446' and spec='FA-128' and StartLine = '2'

if not exists (select * from setting where ctrlno = 'M02446' and spec='FA-128')
insert into Setting
	select 
	   [CtrlNo]
      ,'FA-128'
      ,[LayoutNo]
      ,[Xoffset]
      ,[Yoffset]
      ,[Rotate]
      ,[Current]
      ,[QSW]
      ,[Speed]
      ,[StartLine]
      ,[Block1]
      ,[Block2]
      ,[Block3]
      ,[Block4]
      ,[Block5]
      ,[Block6]
      ,[UseDot]
      ,[UseBlock]
      ,[ComMatrix]
      ,[OptionalSet]  
	from Setting
	where ctrlno = 'M02446' and spec='FA-118'

select * 
from Setting
where ctrlno = 'M02446' and spec='FA-128'

