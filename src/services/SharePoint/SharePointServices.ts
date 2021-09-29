import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { EnvironmentType } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { IListCollection } from "./IList";
import { MockListCollection } from "./data/MockListCollection";
import { MockListItemCollection } from "./data/MockListItemCollection";
import { IListFieldCollection } from "./IListField";
import { IListItemCollection } from "./IListItem";
import { MockListFieldCollection } from "./data/MockListFieldCollection";


export class SharePointServiceManager {
    public context: IWebPartContext;
    public environmentType: EnvironmentType;

    public setup(context: IWebPartContext, environmentType: EnvironmentType): void {
        this.context = context;
        this.environmentType = environmentType;
    }

    public get(relativeEndpointUrl: string): Promise<any> {
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}${relativeEndpointUrl}`, 
        SPHttpClient.configurations.v1).then(response => {
            if(!response.ok) return Promise.reject('GET Request Failed');
            return response.json();
        }).catch(error => {
            return Promise.reject(error);
        });
    }

    public getList(showHiddenLists: boolean = false): Promise<IListCollection>{
        if(this.environmentType == EnvironmentType.Local){
            return new Promise(resolve => resolve(MockListCollection));
        }
        return this.get(`/_api/lists${!showHiddenLists ? '?$filter=Hidden eq flase' : ''}`);
    }

    public getListItmes(listId: string, selectedFields?: string[]): Promise<IListItemCollection> {
        if(this.environmentType == EnvironmentType.Local){
            return new Promise(resolve => resolve(MockListItemCollection));
        }
        return this.get(`/_api/lists/getbyid('${listId}')/items${selectedFields ? `?select=${selectedFields.join(',')}` : ''}`);
    }

    public getListFields(listId: string, showHiddenFields: boolean = false): Promise<IListFieldCollection> {
        if(this.environmentType == EnvironmentType.Local){
            return new Promise(resolve => resolve(MockListFieldCollection));
        }
        return this.get(`/_api/lists/getbyid('${listId}')/fields${!showHiddenFields ? '?$filter=Hidden eq flase' : ''}`);
    }
}

const SharePointService = new SharePointServiceManager();
export default SharePointService;