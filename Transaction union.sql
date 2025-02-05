select * from [dbo].[2019 & 2020 Transactions]
select * from [dbo].[2021 & 2022 transactions]

begin tran
create table Transactions_All(S_N INT, Order_Date varchar(30),TransactionID  varchar(30), CustomerID  varchar(30), 
ProductID  varchar(30), Quantity_Supplied INT, Discount int)

insert into Transactions_All

select * from [dbo].[2019 & 2020 Transactions]
UNION all 
select * from [dbo].[2021 & 2022 transactions]

commit 

select * from Transactions_All