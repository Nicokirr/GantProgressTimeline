import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
const visTimeLine = require("vis/lib/timeline/Timeline");
const visDataSet = require("vis/lib/DataSet")
const moment = require("moment")
const RowRecordId:string = "rowRecId";

export class GantProgressTimeline implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// Cached context object for the latest updateView
	private contextObj: ComponentFramework.Context<IInputs>;
		
	// Div element created as part of this control's main container
	private mainContainer: HTMLDivElement;
	private visuContainer: HTMLDivElement;
	private timeline : any;
	private _notifyOutputChanged: () => void;

	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Need to track container resize so that control could get the available width. The available height won't be provided even this is true
			context.mode.trackContainerResize(true);
			this._notifyOutputChanged = notifyOutputChanged;

			// Create main table container div. 
			this.mainContainer = document.createElement("div");
			this.visuContainer = document.createElement("div");
			this.visuContainer.setAttribute("id","visualisation");
			this.mainContainer.appendChild(this.visuContainer);
			container.appendChild(this.mainContainer);

			

			

	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.contextObj = context;
		if(!context.parameters.dataSetGrid.loading){
				
			// Get sorted columns on View
			let columnsOnView = this.getSortedColumnsOnView(context);

			if (!columnsOnView || columnsOnView.length === 0) {
				return;
			}

			this.createTimeLine(columnsOnView,context.parameters.dataSetGrid);
		}
	}

    private createTimeLine(columnsOnView: DataSetInterfaces.Column[], subgridData: DataSet) {
        const startColumn = 'cr8bb_startdate';
        const endColumn = 'cr8bb_enddate';
        const valueColumn = 'cr8bb_progress';
        const labelColumn = 'cr8bb_description';
        const weightColumn = 'cr8bb_weight';

        if (subgridData.sortedRecordIds.length > 0) {
            var timeLineItems = new visDataSet([]);
            let i = 0;


            for (let currentRecordId of subgridData.sortedRecordIds) {
                let startdate = subgridData.records[currentRecordId].getValue(startColumn);
                let enddate = subgridData.records[currentRecordId].getValue(endColumn);
                let label = subgridData.records[currentRecordId].getFormattedValue(labelColumn);
                let progress = subgridData.records[currentRecordId].getValue(valueColumn);
                let weight = subgridData.records[currentRecordId].getValue(weightColumn);

                if (startdate != null && startdate != "" && label != null && label != "") {
                    let item = {
                        id: ++i,
                        content: label,
                        value: progress,
                        targetvalue: 100 * moment(new Date()).diff(moment(startdate)) / moment(enddate).diff(moment(startdate)),
                        start: moment(startdate).format('YYYY/MM/DD'),
                        end: (enddate != null) ? moment(enddate).format('YYYY/MM/DD') : "",
                        //type : this.DateDiff(enddate,startdate)<1 ? 'point' : 'range',
                        title: "",
                    }
                    item["title"] = (item.end != "") ? item.start + " - " + item.end : item.start;
                    timeLineItems._addItem(item);
                }
            }

            // Configuration for the Timeline
            var options = {
                visibleFrameTemplate: function (item: any) {
                    if (item.visibleFrameTemplate) {
                        return item.visibleFrameTemplate;
                    }
                    let percentage = item.value + '%';
                    let wrapperHTML = '<div class="progress-wrapper">';
                    let progressHTML = '';
                    if (item.value >= item.targetvalue || item.value == 100) {
                        progressHTML = '<div class="progressgood" style="width:' + percentage + '">';
                    } else if (item.value >= item.targetvalue * 0.9) {
                        progressHTML = '<div class="progressmedium" style="width:' + percentage + '">';
                    } else {
                        progressHTML = '<div class="progressbad" style="width:' + percentage + '">';
                    }
                    let progresslabelHTML = '<label class="progress-label" > ' + percentage + ' </label ></div>';
                    return wrapperHTML + progressHTML + progresslabelHTML;
                }
            };

            // Create a Timeline
            if (!this.timeline)
                this.timeline = new visTimeLine(this.visuContainer, timeLineItems, options);
        }
    }
	/**
	* Get sorted columns on view
	* @param context 
	* @return sorted columns object on View
	*/
	private getSortedColumnsOnView(context: ComponentFramework.Context<IInputs>): DataSetInterfaces.Column[]
		{
			if (!context.parameters.dataSetGrid.columns) {
				return [];
			}
			
			let columns =context.parameters.dataSetGrid.columns
				.filter(function (columnItem:DataSetInterfaces.Column) { 
					// some column are supplementary and their order is not > 0
					return columnItem.order >= 0 }
				);
			
			// Sort those columns so that they will be rendered in order
			columns.sort(function (a:DataSetInterfaces.Column, b: DataSetInterfaces.Column) {
				return a.order - b.order;
			});
			
			return columns;
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