-- AUTHOR: Bartosz Caban

-- STORED PROCEDURES

-- #########################################################################################################################

-- The scripts are written and executed on the AdventureWorks2019 database

-- Procedure name: dbo.OrdersAboveThreshold
-- Create a procedure with four parameters that, when called, displays orders above the monetary threshold. The parameters are: @Threshold, @StartYear, @EndYear, @OrderType 

CREATE OR ALTER PROCEDURE dbo.OrdersAboveThreshold(
	@Threshold MONEY
	,@StartYear INT
	,@EndYear INT
	,@OrderType INT
)

AS

BEGIN

	IF @OrderType = 1

		BEGIN 

			SELECT 
				a.SalesOrderID,
				a.OrderDate,
				a.TotalDue

			FROM Sales.SalesOrderHeader a
			INNER JOIN dbo.Calendar b
				ON a.OrderDate = b.DateValue
			WHERE a.TotalDue >= @Threshold 
				AND b.YearNumber BETWEEN @StartYear AND @EndYear
			ORDER BY a.OrderDate

		END

	IF @OrderType = 2 
		
		BEGIN

			SELECT 
				a.PurchaseOrderID,
				a.OrderDate,
				a.TotalDue

			FROM Purchasing.PurchaseOrderHeader a
			INNER JOIN dbo.Calendar b
				ON a.OrderDate = b.DateValue
			WHERE a.TotalDue >= @Threshold 
				AND b.YearNumber BETWEEN @StartYear AND @EndYear
			ORDER BY a.OrderDate

		END

	IF @OrderType = 3

		BEGIN

			SELECT 
				OrderID = a.SalesOrderID,
				a.OrderDate,
				a.TotalDue,
				OrderType = 'Sales'

			FROM Sales.SalesOrderHeader a
			INNER JOIN dbo.Calendar b
				ON a.OrderDate = b.DateValue
			WHERE a.TotalDue >= @Threshold 
				AND b.YearNumber BETWEEN @StartYear AND @EndYear

			UNION ALL

			SELECT 
				a.PurchaseOrderID,
				a.OrderDate,
				a.TotalDue,
				OrderType = 'Purchase'

			FROM Purchasing.PurchaseOrderHeader a
			INNER JOIN dbo.Calendar b
				ON a.OrderDate = b.DateValue
			WHERE a.TotalDue >= @Threshold 
				AND b.YearNumber BETWEEN @StartYear AND @EndYear
			ORDER BY a.OrderDate

		END

END
GO

EXEC dbo.OrdersAboveThreshold 30000, 2010, 2011, 3
GO

-- Procedure name: dbo.OrdersReport
-- Create a procedure with two parameters that, when called, displays the order type: Sales or Purchase. The procedure displays the TopN of the records set in the parameter.
-- The parameters are: @TopN, @OrderType

CREATE OR ALTER PROCEDURE dbo.OrdersReport(
	@TopN INT
	,@OrderType INT
)

AS

BEGIN

	IF @OrderType = 1

		BEGIN
			SELECT
				*
			FROM (
				SELECT 
					ProductName = B.Name,
					LineTotalSum = SUM(A.LineTotal),
					LineTotalSumRank = DENSE_RANK() OVER(ORDER BY SUM(A.LineTotal) DESC)

				FROM AdventureWorks2019.Sales.SalesOrderDetail A
					JOIN AdventureWorks2019.Production.Product B
						ON A.ProductID = B.ProductID

				GROUP BY
					B.Name
				) X

			WHERE LineTotalSumRank <= @TopN
		END

	ELSE

		BEGIN
			SELECT
				*
			FROM (
				SELECT 
					ProductName = B.Name,
					LineTotalSum = SUM(A.LineTotal),
					LineTotalSumRank = DENSE_RANK() OVER(ORDER BY SUM(A.LineTotal) DESC)

				FROM AdventureWorks2019.Purchasing.PurchaseOrderDetail A
					JOIN AdventureWorks2019.Production.Product B
						ON A.ProductID = B.ProductID

				GROUP BY
					B.Name
				) X

			WHERE LineTotalSumRank <= @TopN
		END		

END
GO

EXEC dbo.OrdersReport 15, 2
GO

-- Procedure name: HumanResources.NameSearch
-- Create a procedure with three parameters that, when called, will display the name(s) of the employees according to a pattern given to the parameter.
-- The parameters are: @NameToSearch, @SearchPattern, @MatchType

CREATE OR ALTER PROCEDURE dbo.NameSearch(
	@NameToSearch NVARCHAR(50)
	,@SearchPattern NVARCHAR(100)
	,@MatchType INT
)

AS 

BEGIN

	-- Declaring variables

	DECLARE @sql NVARCHAR(MAX)
	DECLARE @DynamicWhere VARCHAR(MAX)
	DECLARE @NameColumn VARCHAR(100)

	-- Choosing a column to search

	IF @NameToSearch = 'First'
		SET @NameColumn = 'FirstName'

	IF @NameToSearch = 'Middle'
		SET @NameColumn = 'MiddleName'

	IF @NameToSearch = 'Last'
		SET @NameColumn = 'LastName'
	
	-- 1 - exact match
	IF @MatchType = 1
		SET @DynamicWhere = ' = ' + '''' + @SearchPattern + ''''

	-- 2 - begins with
	IF @MatchType = 2
		SET @DynamicWhere = ' LIKE ' + '''' + @SearchPattern  + '%' + ''''

	-- 3 - ends with
	IF @MatchType = 3
		SET @DynamicWhere = ' LIKE ' + '''' + '%' + @SearchPattern + ''''

	-- 4 - contains

	IF @MatchType = 4
		SET @DynamicWhere = ' LIKE ' + '''' + '%' + @SearchPattern + '%' + ''''
	
	-- Executing dynamic sql

	SET @sql = N'SELECT * FROM Person.Person WHERE ' + @NameColumn + @DynamicWhere + ';'
	EXEC(@sql)
	
END
GO

EXEC dbo.NameSearch 'first', 'bar', 4
GO

-- Procedure name: HumanResources.RegisterTime
-- Create an additional table for the HumanResources schema
-- Create a procedure with four parameters. The procedure is designed to record the working time of employees.
-- The parameters are: @BusinessEntityID, @RegIn, @RegOut, @TimeAtWork

CREATE TABLE HumanResources.InOut
(
	Id INT IDENTITY PRIMARY KEY
	,BusinessEntityID INT NOT NULL
		CONSTRAINT FK_InOut_Person_BusinessEntityID FOREIGN KEY REFERENCES Person.BusinessEntity(BusinessEntityID)
	,RegIn DATETIME NULL
	,RegOut DATETIME NULL
);

ALTER TABLE HumanResources.InOut ADD RegInDate AS CAST(RegIn AS DATE) PERSISTED;

CREATE UNIQUE INDEX IX_HumanResources_InOut_BusinessEntityID_RegInDate
ON HumanResources.InOut(BusinessEntityID, RegInDate);
GO

CREATE OR ALTER PROCEDURE HumanResources.RegisterTime (
	@BusinessEntityID INT
	,@RegIn DATETIME = NULL
	,@RegOut DATETIME = NULL
	,@TimeAtWork INT OUTPUT
)

AS

BEGIN
	
	IF @RegIn IS NULL
	BEGIN
		SET @RegIn = GETDATE();
	END

	IF @RegOut IS NULL
	BEGIN
		SET @RegOut = DATEADD(HOUR, 8, @RegIn);
	END

	IF @RegOut IS NOT NULL AND @RegIn > @RegOut
	BEGIN;
		THROW 50002, 'Entry time cannot be later than exit time.', 1;
	END

	IF @RegOut IS NOT NULL AND CAST(@RegIn AS DATE) <> CAST(@RegOut AS DATE)
	BEGIN;
		THROW 50003, 'Entry and exit time must be from the same day.', 1;
	END

	IF NOT EXISTS (SELECT * FROM HumanResources.InOut 
				   WHERE BusinessEntityID = @BusinessEntityID AND RegInDate = CAST(@RegIn AS DATE))
	BEGIN

		INSERT INTO HumanResources.InOut(BusinessEntityID, RegIn, RegOut) 
		VALUES(@BusinessEntityID, @RegIn, @RegOut);

	END

	ELSE
	BEGIN

		UPDATE HumanResources.InOut
		SET RegIn = @RegIn
			,RegOut = @RegOut
		
		WHERE BusinessEntityID = @BusinessEntityID
		AND  RegInDate = CAST(@RegIn AS DATE);		

	END

	SET @TimeAtWork = DATEDIFF(MINUTE, @RegIn, @RegOut)

	RETURN 0
END
GO


DECLARE @TimeAtWork INT
BEGIN TRY
	EXEC HumanResources.RegisterTime 
	13, 
	'2024-10-15 08:00', 
	'2024-10-15 16:10', 
	@TimeAtWork = @TimeAtWork OUTPUT
	SELECT @TimeAtWork AS 'Time at work (min)'
END TRY

BEGIN CATCH
	PRINT 'Operation failed with error: ' + ERROR_MESSAGE()
END CATCH

-- Procedure name: HumanResources.ImportInOut
-- Create a table type SelectedEmployees
-- Create a procedure with two parameters. The procedure is used to import employee entry and exit hour data into the HumanResources.InOut table
-- The parameters are: @InputTable (a table type), @rows

CREATE TYPE SelectedEmployees
AS TABLE (
	BusinessEntityID INT
	,RegIn DATETIME NULL
	,RegOut DATETIME NULL
);
GO

CREATE OR ALTER PROCEDURE HumanResources.ImportInOut (
	@InputTable SelectedEmployees READONLY
	,@rows INT OUTPUT
)

AS 

BEGIN
	
	IF EXISTS (SELECT * FROM @InputTable WHERE RegIn IS NULL OR RegOut IS NULL OR RegIn > RegOut)
	BEGIN;
		THROW 50010, 'Imported data contains empty in or out fields.', 1
	END

	ELSE
	BEGIN
		INSERT INTO HumanResources.InOut (BusinessEntityID, RegIn, RegOut)
		SELECT BusinessEntityID, RegIn, RegOut
		FROM @InputTable

		SET @rows = @@ROWCOUNT;
	END
END

DECLARE @DayInOut SelectedEmployees;
DECLARE @NumOfRows INT 

INSERT INTO @DayInOut (BusinessEntityID, RegIn, RegOut)
VALUES 
(5, '2024-10-16 07:55', '2024-10-16 16:03')
,(2, '2024-10-16 07:53', '2024-10-16 16:05')
,(3, '2024-10-16 07:45', '2024-10-16 16:01')
,(4, '2024-10-16 07:58', '2024-10-16 16:02')

EXECUTE HumanResources.ImportInOut @DayInOut, @NumOfRows OUTPUT
SELECT @NumOfRows AS 'Number of imported rows'

-- Procedure name: dbo.uspGenerateCalendar
-- Create a procedure with two parameters. The procedure generates a calendar.
-- The parameters are: @Year, @Month

CREATE OR ALTER PROCEDURE dbo.uspGenerateCalendar (
	@Year INT
	,@Month INT
)

AS

BEGIN

	SET NOCOUNT ON;

	-- declaring variables

	DECLARE @StartDate DATE = DATEFROMPARTS(@Year, @Month, 1);
	DECLARE @EndDate DATE = EOMONTH(@StartDate);
	DECLARE @DaysInMonth INT = DAY(@EndDate); -- for generating days to loop
	DECLARE @Day INT = 0;
	DECLARE @Calendar TABLE(Day DATE, DayOfWeek NVARCHAR(30));

	-- while loop inserting date into tabular variable

	WHILE @Day < @DaysInMonth
	BEGIN
		INSERT INTO @Calendar(Day)
		VALUES(DATEADD(DAY, @Day, @StartDate));
		SET @Day += 1;
	END;
	
	-- updating the day of the week column in a tabular variable
	
	UPDATE @Calendar SET DayOfWeek = DATEPART(WEEKDAY, DAY);

	SELECT Day, DayOfWeek FROM @Calendar;

END;
GO

-- Create a temporary table and inserting records into table by calling the procedure.

CREATE TABLE #dates (day DATE, day_of_week NVARCHAR(30));
INSERT INTO #dates
EXEC dbo.uspGenerateCalendar 2024, 1;
GO

-- Procedure name: dbo.RegisterDamage
-- Create an additional tables 
-- Create a procedure with two parameters, when called, the procedure enters the damaged products into the DamagedProducts table and subtracts them from the Products table. 
-- The parameters are: @ProductID, @Quantity

CREATE TABLE dbo.Products
(
	ProductID INT IDENTITY PRIMARY KEY
	,ProductName NVARCHAR(50)
	,Quantity INT
);

CREATE TABLE dbo.DamagedProducts
(
	DamagedID INT IDENTITY PRIMARY KEY
	,ProductID INT
	,Quantity INT
);
GO

CREATE OR ALTER PROCEDURE dbo.RegisterDamage (
	@ProductID INT
	,@Quantity INT
)

AS

BEGIN

	SET NOCOUNT ON;

	IF NOT EXISTS (SELECT * FROM dbo.Products WHERE ProductID = @ProductID)
		BEGIN;
			THROW 50019, 'There is no product with this ID.', 1;
		END;

	IF EXISTS (SELECT * FROM dbo.Products WHERE @ProductID = ProductID AND @Quantity > Quantity)
		BEGIN;
			THROW 50020, 'You have exceeded the amount of available products.', 1;
		END;

	ELSE 
		BEGIN TRANSACTION;
			UPDATE dbo.Products 
			SET Quantity = Quantity - @Quantity
			WHERE ProductID = @ProductID;

			INSERT INTO dbo.DamagedProducts(ProductID, Quantity)
			VALUES(@ProductID, @Quantity);
		COMMIT TRANSACTION;

END;
GO

INSERT INTO dbo.Products (ProductName, Quantity)
VALUES 
('Product A', 120),
('Product B', 45),
('Product C', 18);

EXECUTE dbo.RegisterDamage 1, 20

