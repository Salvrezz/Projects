select * from [dbo].[healthcare]

begin tran
UPDATE [dbo].[healthcare]
SET Name = UPPER(LEFT(Name, 1)) + LOWER(SUBSTRING(Name, 2, LEN(Name)));
commit

update [dbo].[healthcare]
SET Gender = CASE 
    WHEN Gender IN ('M', 'm', 'MALE', 'male') THEN 'Male'
    WHEN Gender IN ('F', 'f', 'FEMALE', 'female') THEN 'Female'
    ELSE 'Unknown' -- Handling unexpected values
END;
  
update [dbo].[healthcare]
SET city = UPPER(LEFT(city, 1)) + LOWER(SUBSTRING(city, 2, LEN(city)))