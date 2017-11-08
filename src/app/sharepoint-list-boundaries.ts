import {SharePointList} from './sharepoint-list';
import {SharepSoapWebServices} from './sharep-soap-web-services.service';
import {SubSite} from './subsite';
export class SharePointListBoundaries
{

	constructor(private sharepSoapWebServices: SharepSoapWebServices){}

	getSubSites():Promise<SubSite[]>{
		
		return this.sharepSoapWebServices.getSubSites().then(r=>r);
	}
	
	/** This is the main business method that will run all the logic for the component to show.*/
	getLists():Promise<SharePointList[]>
	{
		let injectedService:SharepSoapWebServices = this.sharepSoapWebServices;
		return injectedService.getSubSites().then
		(			
			function (res)
			{				
				let subsites: SubSite[] = res;	
				let promises : Promise<SharePointList[]> [] = [];
				for(let subsite of subsites)
				{	
					promises.push(injectedService.getListCollectionForSite(subsite.url));
				}
				/*once all the webservices return we can start assembling our table*/
				return Promise.all(promises).then
				(
					function(internalRes)
					{
						let listResults: SharePointList[] = [];
						for(let i of internalRes)
						{ 	
							listResults = listResults.concat(i);
						}
						listResults.sort(function(a,b){return b.itemCount-a.itemCount;});
						return Promise.resolve( listResults);
					}
				)
			}
		).catch(this.handleError);
	}
	
	
	private handleError(error: any): Promise<any> {
	console.error(JSON.stringify(error));
    console.error('Unable to fetch sites.', error); 
    return Promise.reject([]);
  }
}