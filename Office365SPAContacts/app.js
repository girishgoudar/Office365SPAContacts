var o365CorsApp = angular.module("o365CorsApp", ['ngRoute', 'AdalAngular'])
o365CorsApp.factory("ShareData", function () {
    return { value: 0 }
});
o365CorsApp.config(['$routeProvider', '$httpProvider', 'adalAuthenticationServiceProvider', function ($routeProvider, $httpProvider, adalProvider) {
    $routeProvider
           .when('/',
           {
               controller: 'ContactsController',
               templateUrl: 'Contacts.html',
               requireADLogin: true
           })
           .when('/contacts/new',
           {
               controller: 'ContactsNewController',
               templateUrl: 'ContactsNew.html',
               requireADLogin: true
               
           })
           .when('/contacts/edit',
           {
               controller: 'ContactsEditController',
               templateUrl: 'ContactsEdit.html',
               requireADLogin: true

           })
           .when('/contacts/delete',
           {
               controller: 'ContactsDeleteController',
               templateUrl: 'Contacts.html',
               requireADLogin: true

           })
           .otherwise({ redirectTo: '/' });

    var adalConfig = {
        tenant: '5b532de2-3c90-4e6b-bf85-db0ed9cf5b48',
        clientId: '26251bf3-a37f-4cd9-b06e-592e1c1bebf3',
        extraQueryParameter: 'nux=1',
        endpoints: {
           "https://outlook.office365.com/api/v1.0": "https://outlook.office365.com/"
        }
    };
    adalProvider.init(adalConfig, $httpProvider);
}]);
o365CorsApp.controller("ContactsController", function ($scope, $q, $location, $http, ShareData) {
    getContact();
    function getContact() {
        $http.get('https://outlook.office365.com/api/v1.0/me/contacts')
        .then(function (response) {
            $scope.contacts = response.data.value;
        });
    };
    $scope.editUser = function (username) {
        ShareData.value = username;
        $location.path('/contacts/edit');
    };
    $scope.deleteUser = function (contactid) {
        ShareData.value = contactid;
        $location.path('/contacts/delete');
    };
});
o365CorsApp.controller("ContactsNewController", function ($scope, $q,$http,$location) {
    $scope.add = function () {
        var givenName = $scope.givenName
        var surName = $scope.surName
        var email = $scope.email
        contact = { "GivenName": givenName, "Surname": surName, "EmailAddresses": [
             {
                 "Address": email,
                 "Name": givenName
             }
        ]
        }
        
        var options = {
            url: "https://outlook.office365.com/api/v1.0/me/contacts",
            method: 'POST',
            data: JSON.stringify(contact),
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
        };
        return $http(options).then(function (result) {
            $location.path("/#");
        },
        function (error) {
            return error;
        });
    };
});
o365CorsApp.controller("ContactsEditController", function ($scope, $q, $http, $location, ShareData) {
    var id = ShareData.value;
    getContact();
    function getContact() {
        $http.get("https://outlook.office365.com/api/v1.0/me/contacts/"+id)
        .then(function (response) {
            $scope.contact = response.data;
        });
    }
    $scope.update = function () {
        var givenName = $scope.contact.GivenName
        var surName = $scope.contact.Surname
        var email = $scope.contact.EmailAddresses[0].Address
        var id = ShareData.value;

       
        contact = { "GivenName": givenName, "Surname": surName, "EmailAddresses": [
            {
                "Address": email,
                "Name": givenName
            }
            ]
        }
        var options = {
            url: "https://outlook.office365.com/api/v1.0/me/contacts/"+id,
            method: 'PATCH',
            data: JSON.stringify(contact),
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
        };
        return $http(options).then(function (result) {
            $location.path("/");
        },
        function (error) {
            return error;
        });
    };
});
o365CorsApp.controller("ContactsDeleteController", function ($scope, $q, $http, $location, ShareData) {
    var id = ShareData.value;
    getContact();
    function getContact() {
        $http.get("https://outlook.office365.com/api/v1.0/me/contacts/"+id)
        .then(function (response) {
            $scope.contact = response.data;
        });
       
        var id = ShareData.value;
        var options = {
            url: "https://outlook.office365.com/api/v1.0/me/contacts/"+id,
            method: 'DELETE',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
        };

        return $http(options).then(function (result) {
            $location.path("/#");
        },
        function (error) {
            return error;
        });
    }
});
















