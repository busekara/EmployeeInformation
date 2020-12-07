var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
var EmployeeInformationWebPart = /** @class */ (function (_super) {
    __extends(EmployeeInformationWebPart, _super);
    function EmployeeInformationWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.Listname = "EmployeeInformation";
        _this.listItemId = 0;
        return _this;
    }
    EmployeeInformationWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n  <html lang=\"tr\">\n  <head>\n    <!-- Required meta tags -->\n    <meta charset=\"utf-8\">\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1, shrink-to-fit=no\">\n\n    <!-- Bootstrap CSS -->\n    <link rel=\"stylesheet\" href=\"https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css\" integrity=\"sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm\" crossorigin=\"anonymous\"/>\n    \n  </head>\n  <body>\n          <!-- <h1>Personal Information</h1> -->\n      <div class=\"ui-bg-cover ui-bg-overlay-container text-white\" style=\"background:#00BFFF;\">\n        <div class=\"ui-bg-overlay bg-dark opacity-50\"></div>\n        <div class=\"container py-5\">\n          <div class=\"media col-md-10 col-lg-8 col-xl-7 p-0 my-4 mx-auto\">\n            <img src=\"https://arfitect.sharepoint.com/sites/Demo2/_layouts/15/userphoto.aspx?size=L&UserName=" + this.context.pageContext.user.email + "\" alt class=\"d-block ui-w-100 rounded-circle\">\n            <div class=\"media-body ml-5 my-5\">\n              <h4 class=\"font-weight-bold mb-4\">Welcome " + this.context.pageContext.user.displayName + "!</h4>\n            </div>\n          \n          </div>\n            <div class=\"opacity-75>\n                <h5>\n                  <input type=\"text\" class=\"form-control\"  aria-label=\"test\" aria-describedby=\"basic-addon1\" id=\"test\" name=\"fullName\">\n                </h5>\n                <h5>Full Name:\n                  <input type=\"text\" class=\"form-control\"  aria-label=\"fullName\" aria-describedby=\"basic-addon1\" id=\"idFullName\" name=\"fullName\" placeholder=\"Full Name..\" required>\n                </h5>\n                <h5>Address:\n                  <input type=\"text\" class=\"form-control\"  aria-label=\"address\" aria-describedby=\"basic-addon1\" id=\"idAddress\" name=\"address\" placeholder=\"Address..\" required>\n                </h5>\n                <h5>Email:\n                   <input type=\"text\" class=\"form-control\"  aria-label=\"email\" aria-describedby=\"basic-addon1\" id=\"idEmail\"  name=\"emailid\" placeholder=\"Email ID..\" required>\n                </h5>\n                <h5>Mobile:\n                   <input type=\"text\" class=\"form-control\"  aria-label=\"mobile\" aria-describedby=\"basic-addon1\" id=\"idPhoneNumber\" name=\"mobile\" placeholder=\"Mobile Number..\" required>\n                </h5>\n          </div>\n        </div>  \n\n        <td><button class=\"button  btn btn-primary find-Button\"><span> Find </span></button></td>          \n        <td><button class=\"button btn btn-success create-Button\"><span> Create </span></button></td>            \n        <td><button class=\"button btn btn-info update-Button\"><span> Update </span></button></td>            \n        <td><button class=\"button btn btn-danger delete-Button\"><span> Delete </span></button></td>            \n        <td><button class=\"button btn btn-secondary clear-Button\"><span> Clear </span></button></td>\n      </div>\n      <hr />         \n            \n        <table class=\"table table-striped table-hover\" id=\"tblEmployeeInfo\" >\n      </div>\n     ";
        this.setButtonsEventHandlers();
        this.getListData();
    };
    EmployeeInformationWebPart.prototype.setButtonsEventHandlers = function () {
        var webPart = this;
        this.domElement.querySelector('button.find-Button').addEventListener('click', function () { webPart.find(); });
        this.domElement.querySelector('button.create-Button').addEventListener('click', function () { webPart.save(); });
        this.domElement.querySelector('button.update-Button').addEventListener('click', function () { webPart.update(); });
        this.domElement.querySelector('button.delete-Button').addEventListener('click', function () { webPart.delete(); });
        this.domElement.querySelector('button.clear-Button').addEventListener('click', function () { webPart.clear(); });
    };
    EmployeeInformationWebPart.prototype.find = function () {
        var _this = this;
        var emailId = prompt("Enter the Email");
        this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + this.Listname + "')/items?$select=*&$filter=Email eq '" + emailId + "'", SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json()
                .then(function (item) {
                document.getElementById('idFullName')["value"] = item.value[0].Title;
                document.getElementById('idAddress')["value"] = item.value[0].Address;
                document.getElementById('idEmail')["value"] = item.value[0].Email;
                document.getElementById('idPhoneNumber')["value"] = item.value[0].Mobile;
                _this.listItemId = item.value[0].Id;
            });
        });
    };
    EmployeeInformationWebPart.prototype.getListData = function () {
        var _this = this;
        var html = '<table border=1 width=100% style="border-collapse: collapse;">';
        html += "\n    <thead>\n      <tr>\n        <th scope=\"col\">Name</th>\n        <th scope=\"col\">Address</th>\n        <th scope=\"col\">Email</th>\n        <th scope=\"col\">Mobile</th>\n      </tr>\n      </thead>";
        this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + this.Listname + "')/items", SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json()
                .then(function (items) {
                console.log('items.value: ', items.value);
                var listItems = items.value;
                console.log('list items: ', listItems);
                listItems.forEach(function (item) {
                    html += "   \n                 <tr>            \n                 <td  scope=\"col\">" + item.Title + "</td>                  \n                   <td scope=\"col\">" + item.Address + "</td>\n                   <td  scope=\"col\">" + item.Email + "</td>\n                   <td  scope=\"col\">" + item.Mobile + "</td>        \n                 </tr>\n\n                  ";
                });
                html += "</table>\n            ";
                var listContainer = _this.domElement.querySelector('#tblEmployeeInfo');
                listContainer.innerHTML = html;
            });
        });
    };
    EmployeeInformationWebPart.prototype.save = function () {
        var _this = this;
        if (document.getElementById('idFullName')["value"] != "" && document.getElementById('idAddress')["value"] != "" && document.getElementById('idEmail')["value"] != "" && document.getElementById('idPhoneNumber')["value"] != "") {
            var body = JSON.stringify({
                'Title': document.getElementById('idFullName')["value"],
                'Address': document.getElementById('idAddress')["value"],
                'Email': document.getElementById('idEmail')["value"],
                'Mobile': document.getElementById('idPhoneNumber')["value"],
            });
            this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + this.Listname + "')/items", SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'X-HTTP-Method': 'POST'
                },
                body: body
            }).then(function (response) {
                _this.getListData();
                _this.clear();
                alert('Item has been successfully Saved ');
            }, function (error) {
                alert("" + error);
            });
        }
        else {
            alert("Please fill the info!");
        }
    };
    EmployeeInformationWebPart.prototype.update = function () {
        var _this = this;
        if (document.getElementById('idFullName')["value"] != "" && document.getElementById('idAddress')["value"] != "" && document.getElementById('idEmail')["value"] != "" && document.getElementById('idPhoneNumber')["value"] != "") {
            var body = JSON.stringify({
                'Title': document.getElementById('idFullName')["value"],
                'Address': document.getElementById('idAddress')["value"],
                'Email': document.getElementById('idEmail')["value"],
                'Mobile': document.getElementById('idPhoneNumber')["value"],
            });
            this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + this.Listname + "')/items(" + this.listItemId + ")", SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'PATCH'
                },
                body: body
            }).then(function (response) {
                _this.getListData();
                _this.clear();
                alert("Item successfully updated");
            }, function (error) {
                alert("" + error);
            });
        }
        else {
            alert("Please fill the info!");
        }
    };
    EmployeeInformationWebPart.prototype.delete = function () {
        var _this = this;
        if (!window.confirm('Are you sure you want to delete the latest item?')) {
            return;
        }
        this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + this.Listname + "')/items(" + this.listItemId + ")", SPHttpClient.configurations.v1, {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'DELETE'
            }
        }).then(function (response) {
            alert("Item successfully Deleted");
            _this.getListData();
            _this.clear();
        }, function (error) {
            alert("" + error);
        });
    };
    EmployeeInformationWebPart.prototype.clear = function () {
        document.getElementById('idFullName')["value"] = "";
        document.getElementById('idAddress')["value"] = "";
        document.getElementById('idEmail')["value"] = "";
        document.getElementById('idPhoneNumber')["value"] = "";
    };
    return EmployeeInformationWebPart;
}(BaseClientSideWebPart));
export default EmployeeInformationWebPart;
//# sourceMappingURL=EmployeeInformationWebPart.js.map