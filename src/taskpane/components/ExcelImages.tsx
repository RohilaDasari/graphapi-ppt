import { getWorkBooksFromDrive, getWorkSheetsForAWorkBookFromDrive, getChartID, getChartImage } from '../../services/GraphService';
import React = require('react');
import { config } from './Config';

interface CalendarState {
  images: any;
}

export default class ExcelImages extends React.Component<{}, CalendarState> {
  constructor(props) {
    super(props);

    this.state = {
      images: []
    };
  }

  async componentDidMount() {
    try {
      let chartImages = [];
      // Get the user's access token
      var accessToken = await (window.msal as any).acquireTokenSilent({
        scopes: config.scopes
      });
      // Get the user's workbooks in one drive
      var workbooks = await getWorkBooksFromDrive(accessToken);
      for(let workbook of workbooks) {
          let worksheets = await getWorkSheetsForAWorkBookFromDrive(accessToken, workbook.id)
          for(let worksheet of worksheets) {
            let charts = await getChartID(accessToken, workbook.id, worksheet.id)
            for(let chart of charts) {
                let chartImage = await getChartImage(accessToken, workbook.id, worksheet.id, chart.id)
                chartImages.push({
                    workbook: workbook,
                    worksheet: worksheet,
                    chartImage: chartImage
                });
            }
          }
      }
      this.setState({images: chartImages});
    }
    catch(err) {
      console.log(err);
    }
  }

  render() {
    return (
      <div>
        <h1>Images section</h1>
        {this.state.images.map(image => {
            const src = `data:image/png;base64,${image.chartImage}`
            return (
                <>
                <img onClick={() => this._onClick(event, image)} src={src} />
                <br />
                <br />
                </>
            )
        })}
      </div>
    );
  }

  _onClick(_event, image) {
    Office.context.document.setSelectedDataAsync(image.chartImage, {
        coercionType: Office.CoercionType.Image
    }, function (asyncResult) {  
        console.log(asyncResult);             
    });          
  }
}