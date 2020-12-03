Feature: Order Page
	I can login
	I can access Order Page
	I can Load Deal

@mytag
Scenario: Login to Portal, access Order page, Load a deal
	Given I go to portal login page
	Then I provide username, password, deal Urn and login
	And I verify I am logged in
	And I verify the deal is loaded
	Then I verify deal data is accurate compared to Send Order request data
