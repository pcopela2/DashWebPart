import * as React from 'react';
import { IListItem } from '../../../services/SharePoint/IListItem'
import SharePointService from '../../../services/SharePoint/SharePointServices';
import { 
    Bar,
    Line,
    HorizontalBar,
    Pie,
    Doughnut, 
    Radar,
    Polar
} from 'react-chartjs-2';

export interface IChartProps {
    listId: string;
    selectedFields: string[];
    chartType: string;
    chartTitle: string;
    colors: string[];
}

export interface IChartState {
    items: IListItem[];
    loading: boolean;
    error: string | null;
}

export default class Chart extends React.Component<IChartProps, IChartState> {
    constructor(props: IChartProps) {
        super(props);
        
        // Bind methods
        this.getItems = this.getItems.bind(this);
        this.chartData = this.chartData.bind(this);

        //Set initial state
        this.state = {
            items: [],
            loading: false,
            error: null,
        };
    }

    public render(): JSX.Element {
        return (
            <div>
                <h1>{this.props.chartTitle}</h1>

                {this.state.error && <p>{this.state.error}</p>}

                {this.props.chartType == 'Bar' && <Bar data={this.chartData()} {...(arguments as any)}/>}
                {this.props.chartType == 'Line' && <Line data={this.chartData()} {...(arguments as any)}/>}
                {this.props.chartType == 'HorizontalBar' && <HorizontalBar data={this.chartData()} {...(arguments as any)}/>}
                {this.props.chartType == 'Pie' && <Pie data={this.chartData()} {...(arguments as any)}/>}
                {this.props.chartType == 'Doughnut' && <Doughnut data={this.chartData()} {...(arguments as any)}/>}
                {this.props.chartType == 'Radar' && <Radar data={this.chartData()} {...(arguments as any)}/>}
                {this.props.chartType == 'Polar' && <Polar data={this.chartData()} {...(arguments as any)}/>}


                <ul>
                    {this.state.items.map(item => {
                        return(
                            <li key={item.Id}>
                                <strong>{item.Title}</strong> ({item.Id})
                            </li>
                        );
                    })}
                </ul>    

                <button onClick={this.getItems} disabled={this.state.loading}>{this.state.loading ? 'Loading...' : 'Refresh'}</button>
            </div>
        );
    }

    public getItems(): void {
        this.setState({loading: true, items: [], error: null });
        
        SharePointService.getListItmes(this.props.listId).then(items => {
            this.setState({
                items: items.value,
                loading: false,
                error: null,
            });
        }).catch(error => {
            this.setState({
                error: 'Something went wrong.',
                loading: false,
                items: null,
            });
        });
    }

    public chartData(): object {
        //Chart Data
        const data = {
            labels: [],
            datasets: [],
        };

        this.state.items.map((item, i) => {
            const dataset ={
                label: '',
                data: [],
                backgroundColor: this.props.colors[i % this.props.colors.length],
                borderColor: this.props.colors[i % this.props.colors.length],
            };

            //Build dataset
            this.props.selectedFields.map((field, j) => {
                //get the value
                let value = item[field];
                if(value === undefined && item[`OData_${field}`] !== undefined){
                    value = item[`OData_${field}`];
                }

                //Add labels
                if(i == 0 && j > 0){
                    data.labels.push(field);
                }

                if(j == 0){
                    dataset.label = value;
                }else{
                    dataset.data.push(value);
                }
            });

            //Line Chart Check
            if(this.props.chartType == 'Line'){
                dataset['fill'] = false;
            }

            data.datasets.push(dataset);
        });

        return data;
    }
}