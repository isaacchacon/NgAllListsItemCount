export class SharePointList
{
	title:string;
	itemCount:number;
	guid:string;
	parentWeb:string;
	description:string;
	url:string;
	choicesColumns:number;
	/**newly added properties as of 2/6/17***/
	maxItemsPerThrottledOperation:number;
	hiddenProp:boolean;
	enableVersioning:boolean;
	/*** end of newly added properties *****/
	readonly maxChoicesColumns:number = 276;
	currenciesColumns:number;
	readonly maxCurrenciesColumns:number = 72;
	datetimesColumns:number;
	readonly maxDatetimesColumns:number = 48;
	notesColumns:number;
	readonly maxNotesColumns:number = 192;
	numbersColumns:number;
	readonly maxNumbersColumns:number=72;
	usersColumns:number;
	readonly maxUsersColumns:number=16;
	textsColumns: number;
	readonly maxTextsColumns: number= 276;
	booleansColumns: number;
	readonly maxBooleansColumns: number = 96;
	//Will get the size of the whole list.
	
	get Size():number
	{
		return Math.round(((this.choicesColumns*28)+(this.currenciesColumns*12)+(this.datetimesColumns*12)+(this.notesColumns*28)+(this.numbersColumns*48)+(this.usersColumns*4)+(this.textsColumns*28)+(this.booleansColumns*5))*100/7744);
	}
	
	//how close you are to the limit of max. columns.
	get ChoicesPercentage():number
	{
		return Math.round(this.choicesColumns*100/this.maxChoicesColumns);
	}	
	get CurrenciesPercentage():number
	{
		return Math.round(this.currenciesColumns*100/this.maxCurrenciesColumns);
	}	
	get DateTimesPercentage():number
	{
		return Math.round(this.datetimesColumns*100/this.maxDatetimesColumns);
	}	
	get NotesPercentage():number
	{
		return Math.round(this.notesColumns*100/this.maxNotesColumns);
	}	
	get NumbersPercentage():number
	{
		return Math.round(this.numbersColumns*100/this.maxNumbersColumns);
	}	
	get UsersPercentage():number
	{
		return Math.round(this.usersColumns*100/this.maxUsersColumns);
	}	
	get TextsPercentage():number
	{
		return Math.round(this.textsColumns*100/this.maxTextsColumns);
	}	
	get BooleansPercentage():number
	{
		return Math.round(this.booleansColumns*100/this.maxBooleansColumns);
	}	

}