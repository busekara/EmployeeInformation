import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

 export interface IEmployeeInformationWebPartProps {
   description: string;
 }

interface IEmployeeDetails {
  Title:string,
  Address: string;
  Email: string;
  Mobile: string;
}
export default class EmployeeInformationWebPart extends BaseClientSideWebPart<IEmployeeInformationWebPartProps> {
  private Listname: string = "EmployeeInformation";
  private listItemId: number = 0;
  public render(): void {
    this.domElement.innerHTML = `
  <html lang="tr">
  <head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"/>
    
  </head>
  <body>
          <!-- <h1>Personal Information</h1> -->
      <div class="ui-bg-cover ui-bg-overlay-container text-white" style="background:#00BFFF;">
        <div class="ui-bg-overlay bg-dark opacity-50"></div>
        <div class="container py-5">
          <div class="media col-md-10 col-lg-8 col-xl-7 p-0 my-4 mx-auto">
            <img src="https://yoursite.sharepoint.com/sites/yoursite/_layouts/15/userphoto.aspx?size=L&UserName=${this.context.pageContext.user.email}" alt class="d-block ui-w-100 rounded-circle">
            <div class="media-body ml-5 my-5">
              <h4 class="font-weight-bold mb-4">Welcome ${this.context.pageContext.user.displayName}!</h4>
            </div>
          
          </div>
            <div class="opacity-75>
                <h5>
                  <input type="text" class="form-control"  aria-label="test" aria-describedby="basic-addon1" id="test" name="fullName">
                </h5>
                <h5>Full Name:
                  <input type="text" class="form-control"  aria-label="fullName" aria-describedby="basic-addon1" id="idFullName" name="fullName" placeholder="Full Name.." required>
                </h5>
                <h5>Address:
                  <input type="text" class="form-control"  aria-label="address" aria-describedby="basic-addon1" id="idAddress" name="address" placeholder="Address.." required>
                </h5>
                <h5>Email:
                   <input type="text" class="form-control"  aria-label="email" aria-describedby="basic-addon1" id="idEmail"  name="emailid" placeholder="Email ID.." required>
                </h5>
                <h5>Mobile:
                   <input type="text" class="form-control"  aria-label="mobile" aria-describedby="basic-addon1" id="idPhoneNumber" name="mobile" placeholder="Mobile Number.." required>
                </h5>
          </div>
        </div>  

        <td><button class="button  btn btn-primary find-Button"><span> Find </span></button></td>          
        <td><button class="button btn btn-success create-Button"><span> Create </span></button></td>            
        <td><button class="button btn btn-info update-Button"><span> Update </span></button></td>            
        <td><button class="button btn btn-danger delete-Button"><span> Delete </span></button></td>            
        <td><button class="button btn btn-secondary clear-Button"><span> Clear </span></button></td>
      </div>
      <hr />         
            
        <table class="table table-striped table-hover" id="tblEmployeeInfo" >
      </div>
     `;
    this.setButtonsEventHandlers();
    this.getListData();
  }
 
  private setButtonsEventHandlers(): void {
    const webPart: EmployeeInformationWebPart = this;
    this.domElement.querySelector('button.find-Button').addEventListener('click', () => { webPart.find(); });
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.save(); });
    this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.update(); });
    this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart.delete(); });
    this.domElement.querySelector('button.clear-Button').addEventListener('click', () => { webPart.clear(); });
  }
 
  private find(): void {
    let emailId = prompt("Enter the Email");
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items?$select=*&$filter=Email eq '${emailId}'`, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json()
          .then((item: any): void => {
            document.getElementById('idFullName')["value"] = item.value[0].Title;
            document.getElementById('idAddress')["value"] = item.value[0].Address;
            document.getElementById('idEmail')["value"] = item.value[0].Email;
            document.getElementById('idPhoneNumber')["value"] = item.value[0].Mobile;
            this.listItemId = item.value[0].Id;
          });
      });
  }
 
  private getListData() {
    let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
    html += `
    <thead>
      <tr>
        <th scope="col">Name</th>
        <th scope="col">Address</th>
        <th scope="col">Email</th>
        <th scope="col">Mobile</th>
      </tr>
      </thead>`
    ;
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items`, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json()
          .then((items: any): void => {
            console.log('items.value: ', items.value);
            let listItems: IEmployeeDetails[] = items.value;
            console.log('list items: ', listItems);
 
            listItems.forEach((item: IEmployeeDetails) => {
              html += `   
                 <tr>            
                 <td  scope="col">${item.Title}</td>                  
                   <td scope="col">${item.Address}</td>
                   <td  scope="col">${item.Email}</td>
                   <td  scope="col">${item.Mobile}</td>        
                 </tr>

                  `;
            });
            html += `</table>
            `;
            const listContainer: Element = this.domElement.querySelector('#tblEmployeeInfo');
            listContainer.innerHTML = html;
          });
      });
  }
 
  private save(): void {
    if(document.getElementById('idFullName')["value"]!="" && document.getElementById('idAddress')["value"]!="" && document.getElementById('idEmail')["value"]!="" && document.getElementById('idPhoneNumber')["value"]!=""){
      const body: string = JSON.stringify({
        'Title': document.getElementById('idFullName')["value"],
        'Address': document.getElementById('idAddress')["value"],
        'Email': document.getElementById('idEmail')["value"],
        'Mobile': document.getElementById('idPhoneNumber')["value"],
      });
   
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'POST'
          },
          body: body
        }).then((response: SPHttpClientResponse): void => {
          this.getListData();
          this.clear();
          alert('Item has been successfully Saved ');
        }, (error: any): void => {
          alert(`${error}`);
        });

    }
    else{alert("Please fill the info!");}
    
  }
 
  private update(): void {
    if(document.getElementById('idFullName')["value"]!="" && document.getElementById('idAddress')["value"]!="" && document.getElementById('idEmail')["value"]!="" && document.getElementById('idPhoneNumber')["value"]!=""){

      const body: string = JSON.stringify({
        'Title': document.getElementById('idFullName')["value"],
        'Address': document.getElementById('idAddress')["value"],
        'Email': document.getElementById('idEmail')["value"],
        'Mobile': document.getElementById('idPhoneNumber')["value"],
      });
  
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${this.listItemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'PATCH'
          },
          body: body
        }).then((response: SPHttpClientResponse): void => {
          this.getListData();
          this.clear();
          alert(`Item successfully updated`);
        }, (error: any): void => {
          alert(`${error}`);
        });
    }
    else{alert("Please fill the info!");}

  }
 
  private delete(): void {
    if (!window.confirm('Are you sure you want to delete the latest item?')) {
      return;
    }
 
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${this.listItemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        }
      }).then((response: SPHttpClientResponse): void => {
        alert(`Item successfully Deleted`);
        this.getListData();
        this.clear();
      }, (error: any): void => {
        alert(`${error}`);
      });
  }
 
  private clear(): void {
    document.getElementById('idFullName')["value"] = "";
    document.getElementById('idAddress')["value"] = "";
    document.getElementById('idEmail')["value"] = "";
    document.getElementById('idPhoneNumber')["value"] = "";
  }
 
}
