import { Component } from '@angular/core';
import { OnInit } from '@angular/core';
import {SubSite} from './subsite';
import {SharePointList} from './sharepoint-list';
import {SharepSoapWebServices} from './sharep-soap-web-services.service';
import {SharePointListBoundaries} from './sharepoint-list-boundaries';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styles: [`
    .even { background-color: #F3EBDC; }
    .odd { background-color: #EFE3CE; }
	a {cursor:pointer;]
    `],
	providers:[SharepSoapWebServices]
})
export class AppComponent {
text: string = 'Live Count of all items in all lists and libraries';
  title = 'app';
  subsites: SubSite[];
	lists: SharePointList[];
	sortOrderAscending:boolean = true;
	
	constructor(private sharepSoapbWebServices: SharepSoapWebServices){}
	ngOnInit():void{
		this.getSubSites();
		this.getLists();
	}
	
	getListStyle(list:SharePointList):any{
		/*** this is over the limit ***/
		if((list.itemCount/1) > (list.maxItemsPerThrottledOperation/1))
		{
			return {'background-color':'#ff0000b3', 'text-align':'right'};
		}
		/*** we are giving a warning at 80% ***/
		if((list.itemCount/1) > (list.maxItemsPerThrottledOperation*.8))
		{
			return {'background-color':'#ebff00cc','text-align':'right'};
		}
		return {'background-color':'transparent','text-align':'right'};
	}
	
	getSubSites():void{
		 let boundaries = new SharePointListBoundaries(this.sharepSoapbWebServices);
		boundaries.getSubSites().then(r => this.subsites=r);
	}
	getLists():void{
		 let boundaries = new SharePointListBoundaries(this.sharepSoapbWebServices);
		 boundaries.getLists().then(r=> this.lists = r);
	}
	sort(column:string, asc:boolean):void{
		this.lists.sort(function(a,b)
		{
			if(asc)
			{
				if(column=='itemCount')
				{
					return a.itemCount - b.itemCount;
				}
				if(column == 'title')
				{
					  //compare two values
					if(a.title.toLowerCase() < b.title.toLowerCase()) return -1;
					if(a.title.toLowerCase() > b.title.toLowerCase()) return 1;
					return 0;
				}
				if(column == 'parentWeb')
				{
				  //compare two values
					if(a.parentWeb.toLowerCase() < b.parentWeb.toLowerCase()) return -1;
					if(a.parentWeb.toLowerCase() > b.parentWeb.toLowerCase()) return 1;
					return 0;
				}
				
				if(column == 'hiddenProp')
				{
				  //compare two values
					if(a.hiddenProp && !b.hiddenProp) return -1;
					if(!a.hiddenProp && b.hiddenProp) return 1;
					return 0;
				}
				if(column == 'enableVersioning')
				{
				  //compare two values
					if(a.enableVersioning && !b.enableVersioning) return -1;
					if(!a.enableVersioning && b.enableVersioning) return 1;
					return 0;
				}
			}
			else
			{	
				if(column=='itemCount')
				{
					return b.itemCount - a.itemCount;
				}
				if(column == 'title')
				{
					  //compare two values
					if(a.title.toLowerCase() > b.title.toLowerCase()) return -1;
					if(a.title.toLowerCase() < b.title.toLowerCase()) return 1;
					return 0;
				}
				if(column == 'parentWeb')
				{
				  //compare two values
					if(a.parentWeb.toLowerCase() > b.parentWeb.toLowerCase()) return -1;
					if(a.parentWeb.toLowerCase() < b.parentWeb.toLowerCase()) return 1;
					return 0;
				}
				
				if(column == 'hiddenProp')
				{
				  //compare two values
					if(!a.hiddenProp && b.hiddenProp) return -1;
					if(a.hiddenProp && !b.hiddenProp) return 1;
					return 0;
				}
				if(column == 'enableVersioning')
				{
				  //compare two values
					if(!a.enableVersioning && b.enableVersioning) return -1;
					if(a.enableVersioning && !b.enableVersioning) return 1;
					return 0;
				}
			}
		});
	}
  
}
