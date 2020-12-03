Feature: IntegrationRequests
	Integration Login POST request
	Integration Send Order Request

Scenario: POST Integration Login Request, Verify, Store generated Token
	Given I make call to integration login POST request
	Then I verify login response is valid
	And I store token from response
	And Create and write Login request and response data to excel file

Scenario: POST Integration Send Order Request with generated Token from Login Request and Request Body, Verify, Store generated DealURN
	Given I make call to integration send order POST request
	Then I verify send order response is valid
	And I store dealURN from response
	And Create and write Send Order request and response data to excel file
