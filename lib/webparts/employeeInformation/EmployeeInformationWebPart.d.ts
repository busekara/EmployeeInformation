import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IEmployeeInformationWebPartProps {
    description: string;
}
export interface IRestApiDemoWebPartProps {
    description: string;
}
export default class EmployeeInformationWebPart extends BaseClientSideWebPart<IEmployeeInformationWebPartProps> {
    private Listname;
    private listItemId;
    render(): void;
    private setButtonsEventHandlers;
    private find;
    private getListData;
    private save;
    private update;
    private delete;
    private clear;
}
//# sourceMappingURL=EmployeeInformationWebPart.d.ts.map