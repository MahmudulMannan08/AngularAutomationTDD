Feature: IntegrationReqDataDriven
	Integration Login POST request
	Read Data From Excel
	Integration Send Order Request

@ignore
Scenario: POST Integration Login Request, Verify, Store generated Token
	Given I make call to integration login POST request
	Then I verify login response is valid
	And I store token from response
	And Create and write Login request and response data to excel file
@ignore
Scenario Outline: POST Integration Send Order Request with generated Token from Login Request and Reading Request Body Data from excel, Verify, Store generated DealURN
	#Given I read Send Order request data from excel
	Given I read Send Order request data <Cell Start>, <Cell End> & <Sheet Number> from excel
	Then I make call to integration send order POST request with data from excel
	Then I verify send order response is valid
	And I store dealURN from response
	And Create and write Send Order request and response data to excel file

Examples: 
	| Cell Start | Cell End | Sheet Number |
	| C4         | BJ4      | 1            |
	| C5         | BJ5      | 1            |
	| C6         | BJ6      | 1            |
	| C7         | BJ7      | 1            |
	| C8         | BJ8      | 1            |
	| C9         | BJ9      | 1            |
	| C10        | BJ10     | 1            |
	| C11        | BJ11     | 1            |
	| C12        | BJ12     | 1            |
	| C13        | BJ13     | 1            |
	| C14        | BJ14     | 1            |
	| C15        | BJ15     | 1            |
	| C16        | BJ16     | 1            |
	| C17        | BJ17     | 1            |
	| C18        | BJ18     | 1            |
	| C19        | BJ19     | 1            |
	| C20        | BJ20     | 1            |
	| C21        | BJ21     | 1            |
	| C22        | BJ22     | 1            |
	| C23        | BJ23     | 1            |
