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
        clientId: '2b775a7e-0f86-45a6-9757-5f1903db9fd6',
        extraQueryParameter: 'nux=1',
        endpoints: {
           "https://outlook.office365.com/api/v1.0": "https://outlook.office365.com/"
        }
    };
    adalProvider.init(adalConfig, $httpProvider);
}]);
o365CorsApp.controller("ContactsController", function ($scope, $q, $location, $http, ShareData, o365CorsFactory) {
    o365CorsFactory.getContacts().then(function (response) {
        $scope.contacts = response.data.value;
    });

    $scope.editUser = function (userName) {
        ShareData.value = userName;
        $location.path('/contacts/edit');
    };
    $scope.deleteUser = function (contactId) {
        ShareData.value = contactId;
        $location.path('/contacts/delete');
    };
});
o365CorsApp.controller("ContactsNewController", function ($scope, $q, $http, $location, o365CorsFactory) {
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

    o365CorsFactory.insertContact().then(function (contact) {
        $location.path("/#");
    });
});
o365CorsApp.controller("ContactsEditController", function ($scope, $q, $http, $location, ShareData, o365CorsFactory) {
    var id = ShareData.value;
    //getContact();
    //function getContact() {
    //    $http.get("https://outlook.office365.com/api/v1.0/me/contacts/"+id)
    //    .then(function (response) {
    //        $scope.contact = response.data;
    //    });
    //}
    //alert(id);
    o365CorsFactory.getContact(id).then(function (response) {
        $scope.contact = response.data;
    });

    var givenName = $scope.contact.GivenName
    //var surName = $scope.contact.Surname
    //var email = $scope.contact.EmailAddresses[0].Address
   

    //contact = {
    //    "GivenName": givenName, "Surname": surName, "EmailAddresses": [
    //            {
    //                "Address": email,
    //                "Name": givenName
    //            }
    //    ]
    //}

    //o365CorsFactory.updateContact(contact, id).then(function () {
    //    $location.path("/#");
    //});

    //$scope.update = function () {
    //    var options = {
    //        url: "https://outlook.office365.com/api/v1.0/me/contacts/"+id,
    //        method: 'PATCH',
    //        data: JSON.stringify(contact),
    //        headers: {
    //            'Accept': 'application/json',
    //            'Content-Type': 'application/json'
    //        },
    //    };
    //    return $http(options).then(function (result) {
    //        $location.path("/");
    //    },
    //    function (error) {
    //        return error;
    //    });
    //};
});
o365CorsApp.controller("ContactsDeleteController", function ($scope, $q, $http, $location, ShareData, o365CorsFactory) {
    var id = ShareData.value;
    //deleteContact();
    //function deleteContact() {
    //   var options = {
    //        url: "https://outlook.office365.com/api/v1.0/me/contacts/"+id,
    //        method: 'DELETE',
    //        headers: {
    //            'Accept': 'application/json',
    //            'Content-Type': 'application/json'
    //        },
    //    };

    o365CorsFactory.deleteContact().then(function (id) {
        $location.path("/#");
    });

    //    return $http(options).then(function (result) {
    //        $location.path("/#");
    //    },
    //    function (error) {
    //        return error;
    //    });
    //}
});

o365CorsApp.factory('o365CorsFactory', ['$http', function ($http) {
    var factory = {};
   
    factory.getContacts = function () {
        return $http.get('https://outlook.office365.com/api/v1.0/me/contacts')
    }

    factory.getContact = function (id) {
        return $http.get('https://outlook.office365.com/api/v1.0/me/contacts/'+id)
    }

    factory.insertContact = function (contact) {
        var options = {
            url: 'https://outlook.office365.com/api/v1.0/me/contacts',
            method: 'POST',
            data: JSON.stringify(contact),
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
        };
    }

    factory.updateContact = function (contact,id) {
        return options = {
            url: 'https://outlook.office365.com/api/v1.0/me/contacts/'+id,
            method: 'PATCH',
            data: JSON.stringify(contact),
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
        };
    }

    factory.deleteContact = function (id) {
        return options = {
            url: 'https://outlook.office365.com/api/v1.0/me/contacts/'+id,
            method: 'DELETE',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            },
        };
        
    }

    return factory;
}]);

















