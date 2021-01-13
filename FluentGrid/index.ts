import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import {IDataSetProps,HoverCardBasicExample} from "./reactGrid"
import { DetailsList, buildColumns, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as ReactDOM from "react-dom";
import * as React from "react";
import {initializeIcons} from '@uifabric/icons';

type DataSet = ComponentFramework.PropertyTypes.DataSet;
//this._context.parameters.dataSetGrid.paging.loadNextPage()

class whoAmIRequest{
	constructor(){}
}
interface whoAmIRequest {
	getMetadata():any;
	};

export class FluentGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _props: IDataSetProps;
	private _context: ComponentFramework.Context<IInputs>;
	private _inputElement: React.ReactElement;
	private _container: HTMLDivElement;
	private _gridContainer: HTMLDivElement;
	private _notifyOutputChanged: () => void;
	private _array1: object[];
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	private buttonClick(): void{


	
	// Construct a request object from the metadata

	whoAmIRequest.prototype.getMetadata = function(){
			return	{
				boundParameter: null,
				parameterTypes: {},
				operationType: 1, // This is a function. Use '0' for actions and '2' for CRUD
				operationName: "WhoAmI"
				};
	};

		this._context.webAPI.execute(new whoAmIRequest()).then(
			function (response) {
				if (response.ok) {
					console.log("Status: %s %s", response.status, response.statusText);
		
					// Use response.json() to access the content of the response body.
					response.json().then(
						function (responseBody:any) {
							console.log("User Id: %s", responseBody.UserId);
							alert(responseBody.UserId);
							// perform other operations as required;
						});
				}
			},
			function (error) {
				console.log(error.message);
				// handle error conditions
			}
		);
	}
	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		//https://developer.microsoft.com/en-us/fluentui#/styles/web/icons#fabric-react
		//do this right away any time you are using fluentui icons
		initializeIcons();
		// Add control initialization code
		this._context = context;
		this._container = container;
		context.mode.trackContainerResize(true);
        // Create main table container div. 
        this._container = document.createElement("div");
        // Create data table container div. 
		this._gridContainer = document.createElement("div");
		//styling added here from trial and error, if removed, scroll bars will not show up on the grid list
		this._gridContainer.classList.add("DataSetControl_grid-container");
		this._gridContainer.setAttribute("style","height:99%;position: inherit;overflow:scroll");
        this._container.appendChild(this._gridContainer);
		this._container.classList.add("DataSetControl_main-container");
		this._container.setAttribute("style","height:99%;position: inherit;overflow:scroll");
		container.appendChild(this._container);
		this._notifyOutputChanged = notifyOutputChanged;

		//the dataset is provided by pcf in a complex object format that contains the full columnset in each record 
		//as well as a number of other properties that aren't needed every time
		//in order to provide a meaningful, barebones array of data to the react list control
		let fullEntityRecords:ComponentFramework.PropertyHelper.DataSetApi.EntityRecord[] = [];
		let columnsOnView:IColumn[] = [];
		let formattedDataSet:Object[] = [];
		let array3 = Object.values(context.parameters.sampleDataSet ? context.parameters.sampleDataSet.records:{});
		let cols = Object.values(array3[0]._getColumns()).map(function(value){
			columnsOnView.push({
				fieldName:value.name,
				key:value.name,
				name:value.displayName,
				minWidth:100,
				maxWidth:200,
				isCollapsible:true,
				isCollapsable:true,
				isGrouped:false,
				isMultiline:false,
				isResizable:true,
				isRowHeader:false,
				isSorted:false,
				isSortedDescending:false,
				columnActionsMode:1
			})
		});
		let array2 = Object.values(context.parameters.sampleDataSet ? context.parameters.sampleDataSet.records: {}).map(function(value)
		{
			fullEntityRecords.push(value);
		});
		for(var rec of fullEntityRecords){
			let  o2:any = {};
			for(var col of columnsOnView){
				if(rec._record.fields[col.key]){
					let o1 = rec._record.fields[col.key].value?rec._record.fields[col.key].value:"";
					o2[col.key] = o1;
				}
			}
			formattedDataSet.push(o2);
		}
		this._props = {
			cols:columnsOnView,
			data:formattedDataSet,
			onButtonClicked: this.buttonClick.bind(this)
		}
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
		ReactDOM.render(
			this._inputElement = React.createElement(HoverCardBasicExample,this._props),this._container
		);
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

}
