import {Injectable} from '@angular/core';
import {Headers, Http} from '@angular/http';

import 'rxjs/add/operator/toPromise';

import {SubSite} from './subsite';
import {SUBSITES} from './mock-subsites';
import {SharePointList} from './sharepoint-list';
declare var $:any;

@Injectable()
export class SharepSoapWebServices{
	private mockResponse2='xml2';
  private websSubSiteUrl = '/_vti_bin/Webs.asmx';  // URL to subsite
  private listsCollectionUrl = '/_vti_bin/Lists.asmx'
  private subSiteOperation = 'GetAllSubWebCollection';//web service name.
  private listsCollectionOperation = 'GetListCollection';
  private websPayload = `<?xml version="1.0" encoding="utf-8"?>
	<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
		<soap12:Body>
			<GetAllSubWebCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/" />
		</soap12:Body>
	</soap12:Envelope>`;
	private listsCollectionPayload = `<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>
    <GetListCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/" />
  </soap12:Body>
</soap12:Envelope>`

  constructor(private http: Http) { }
  getSubSites(): Promise<SubSite[]>
  {		
  /*
	let filledResponse: SubSite[]   = [];
				
				$(this.mockResponse).find('Web').each(function( index:any ) {
					filledResponse.push({title:$(this).attr('Title'),url:$(this).attr('Url')}); 
			});
				return Promise.resolve( filledResponse);
	*/
  
  	return this.http.post(this.websSubSiteUrl, this.websPayload,{headers:this.headers,})
	.toPromise()
	.then( function(res)
			{
			 let filledResponse: SubSite[]   = [];
				$(res.text()).find('Web').each(function( index:any ) {
					filledResponse.push({title:$(this).attr('Title'),url:$(this).attr('Url')}); 
			});
				return filledResponse;
			})
	.catch(this.handleError);
	
  }
   
  getListCollectionForSite(subSiteUrl:string):Promise<SharePointList[]>
  {
	return this.http.post(subSiteUrl+this.listsCollectionUrl, this.listsCollectionPayload,{headers:this.headers,})
	.toPromise()
	.then( function(res)
			{
			let filledResponse: SharePointList[]   = [];
				$(res.text()).find('List').each(function( index:any ) {
					let tempList:SharePointList = new SharePointList();
					tempList.guid=$(this).attr('ID');
					tempList.title=$(this).attr('Title');
					tempList.itemCount = $(this).attr('ItemCount');
					tempList.description = $(this).attr('Description');
					tempList.url = $(this).attr('DefaultViewUrl');
					tempList.parentWeb= $(this).attr('WebFullUrl');
					tempList.maxItemsPerThrottledOperation = $(this).attr('MaxItemsPerThrottledOperation');
					tempList.hiddenProp = (this.attributes[62].value=="True");
					tempList.enableVersioning = ($(this).attr('EnableVersioning')=="True");
					filledResponse.push(tempList); 
			});
				return filledResponse;
			})
	.catch(this.handleError);
  }
  
  
  
  
  
  
  
  private handleError(error: any): Promise<any> {
	
    console.error('An error occurred', error); // for demo purposes only
    return Promise.reject(error.message || error);
  }

  private headers = new Headers({
								 'Content-Type': 'application/soap+xml; charset=utf-8',
								 });
	
	private mockResponse =`<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
   <soap:Body>
      <GetAllSubWebCollectionResponse xmlns="http://schemas.microsoft.com/sharepoint/soap/">
         <GetAllSubWebCollectionResult>
            <Webs>
               <Web Title="Home" Url="https://sp2010-tax.sp.ohio.gov" />
               <Web Title="Adoption" Url="https://sp2010-tax.sp.ohio.gov/Adoption" />
               <Web Title="Apps" Url="https://sp2010-tax.sp.ohio.gov/apps" />
               <Web Title="AuditTools" Url="https://sp2010-tax.sp.ohio.gov/apps/AuditTools" />
               <Web Title="Department Policies" Url="https://sp2010-tax.sp.ohio.gov/apps/AuditTools/DeptPol" />
               <Web Title="Bad Check" Url="https://sp2010-tax.sp .ohio.gov/apps/BadCheck" />
               <Web Title="BillTrac Repository" Url="https://sp2010-tax.sp.ohio.gov/apps /BillTrac" />
               <Web Title="CART" Url="https://sp2010-tax.sp.ohio.gov/apps/CART" />
               <Web Title="Contract  Repository" Url="https://sp2010-tax.sp.ohio.gov/apps/ContractRepo" />
               <Web Title="RefundHandling" Url="https://sp2010-tax.sp.ohio.gov/apps/RefundHandling" />
               <Web Title="SART" Url="https://sp2010-tax.sp.ohio.gov/apps/SART" />
               <Web Title="SeparationChecklist" Url="https://sp2010-tax.sp.ohio.gov/apps/SeparationChecklist" />
               <Web Title="Tax Forms" Url="https://sp2010-tax.sp.ohio.gov/apps/taxforms" />
               <Web Title="Taxipedia" Url="https://sp2010-tax.sp.ohio.gov/apps/Taxipedia" />
               <Web Title="Referrals to Compliance" Url="https://sp2010-tax.sp.ohio.gov/apps/TPSReferrals" />
               <Web Title="Audit" Url="https://sp2010-tax.sp.ohio.gov /Audit" />
               <Web Title="Forms" Url="https://sp2010-tax.sp.ohio.gov/Forms" />
               <Web Title="4549 Return" Url="https://sp2010-tax.sp.ohio.gov/Forms/4549Ret" />
               <Web Title="AG Memo" Url="https://sp2010-tax.sp.ohio .gov/Forms/AGMemo" />
               <Web Title="Asset Verification" Url="https://sp2010-tax.sp.ohio.gov/Forms/AssetVerify" />
               <Web Title="Biennium Budget" Url="https://sp2010-tax.sp.ohio.gov/Forms/BienniumBudget" />
               <Web Title="Check Handling" Url="https://sp2010-tax.sp.ohio.gov/Forms/ChkHndlng" />
               <Web Title="Change of Address  - Form ADM 4058" Url="https://sp2010-tax.sp.ohio.gov/Forms/COA" />
               <Web Title="Conflict Of Interest" Url="https://sp2010-tax.sp.ohio.gov/Forms/Conflict Of Interest" />
               <Web Title="Contracts" Url="https ://sp2010-tax.sp.ohio.gov/Forms/Contracts" />
               <Web Title="Electronic Sign-Off" Url="https://sp2010-tax .sp.ohio.gov/Forms/ESignOff" />
               <Web Title="IDBadge" Url="https://sp2010-tax.sp.ohio.gov/Forms/IDBadge" />
               <Web Title="IncidentReport" Url="https://sp2010-tax.sp.ohio.gov/Forms/IR" />
               <Web Title="ISD Portal" Url="https://sp2010-tax.sp.ohio.gov/Forms/ISDPortal" />
               <Web Title="Individual Teleworking Agreement" Url="https://sp2010-tax.sp.ohio.gov/Forms/ITA" />
               <Web Title="Lost or Stolen Equipment" Url="https:/ /sp2010-tax.sp.ohio.gov/Forms/LostStolenEquipmentv2" />
               <Web Title="Payment Card Log" Url="https://sp2010-tax .sp.ohio.gov/Forms/PCL" />
               <Web Title="PIT Referrals" Url="https://sp2010-tax.sp.ohio.gov/Forms/PITRef" />
               <Web Title="Project Survey" Url="https://sp2010-tax.sp.ohio.gov/Forms/ProjectSurvey" />
               <Web Title="Records Request" Url="https://sp2010-tax.sp.ohio.gov/Forms/RR" />
               <Web Title="Reports To Change" Url="https://sp2010-tax.sp.ohio.gov/Forms/RTC" />
               <Web Title="Request to Purchase" Url="https://sp2010-tax .sp.ohio.gov/Forms/RTP" />
               <Web Title="Safety Warden Drill" Url="https://sp2010-tax.sp.ohio.gov/Forms /Safety Warden Drill" />
               <Web Title="Suspicious Filer Referral" Url="https://sp2010-tax.sp.ohio.gov/Forms /SFR" />
               <Web Title="SpeakingEngagement" Url="https://sp2010-tax.sp.ohio.gov/Forms/SpeakingEngagement" />
               <Web Title="Supply Request" Url="https://sp2010-tax.sp.ohio.gov/Forms/Supply" />
               <Web Title="TaxForms" Url="https://sp2010-tax.sp.ohio.gov/Forms/TaxForms" />
               <Web Title="TestDocCenter" Url="https://sp2010-tax .sp.ohio.gov/Forms/TaxForms/TestDocCenter" />
               <Web Title="TaxTAP" Url="https://sp2010-tax.sp.ohio.gov /Forms/TaxTAP3" />
               <Web Title="TOMAS" Url="https://sp2010-tax.sp.ohio.gov/Forms/tomas" />
               <Web Title="Training  Request Form" Url="https://sp2010-tax.sp.ohio.gov/Forms/TRF" />
               <Web Title="Vehicle Registration" Url="https://sp2010-tax.sp.ohio.gov/Forms/VehicleRegistration" />
               <Web Title="Withholding Exemption Certificate" Url="https://sp2010-tax.sp.ohio.gov/Forms/WEC" />
               <Web Title="Wireless Telecommunications Request" Url="https://sp2010-tax.sp.ohio.gov/Forms/WTR" />
               <Web Title="HR" Url="https://sp2010-tax.sp.ohio.gov/HR" />
               <Web Title="IS" Url="https://sp2010-tax.sp.ohio.gov/IS" />
               <Web Title="Data Capture" Url="https://sp2010-tax .sp.ohio.gov/IS/dc" />
               <Web Title="Network" Url="https://sp2010-tax.sp.ohio.gov/IS/Network" />
               <Web Title="PMO" Url="https://sp2010-tax.sp.ohio.gov/IS/PMO" />
               <Web Title="Web Services" Url="https://sp2010-tax .sp.ohio.gov/IS/webservices" />
               <Web Title="Legislation" Url="https://sp2010-tax.sp.ohio.gov/legislation" />
               <Web Title="NotesMigrationTestData" Url="https://sp2010-tax.sp.ohio.gov/nmsptest" />
               <Web Title="NotesMigration" Url="https://sp2010-tax.sp.ohio.gov/NotesMigration" />
               <Web Title="Organizational Development" Url="https ://sp2010-tax.sp.ohio.gov/OrgD" />
               <Web Title="STARS Training" Url="https://sp2010-tax.sp.ohio.gov/OrgD /STARSTraining" />
               <Web Title="Projects" Url="https://sp2010-tax.sp.ohio.gov/projects" />
               <Web Title="Business  Modernization (BMOD)" Url="https://sp2010-tax.sp.ohio.gov/projects/BMOD" />
               <Web Title="SCCM Windows  10" Url="https://sp2010-tax.sp.ohio.gov/Sccm" />
               <Web Title="SOP" Url="https://sp2010-tax.sp.ohio.gov /SOP" />
               <Web Title="Test" Url="https://sp2010-tax.sp.ohio.gov/Test" />
               <Web Title="ID Migration" Url="https ://sp2010-tax.sp.ohio.gov/Test/IdmIGRATION" />
               <Web Title="Taxpayer Services" Url="https://sp2010-tax .sp.ohio.gov/TPS" />
            </Webs>
         </GetAllSubWebCollectionResult>
      </GetAllSubWebCollectionResponse>
   </soap:Body>
</soap:Envelope>`;
}