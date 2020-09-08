import { Table } from 'reactstrap';
import { getEvents } from '../../services/GraphService';
import React = require('react');
import { config } from './Config';

interface CalendarState {
  events: any;
}

export default class Calendar extends React.Component<{}, CalendarState> {
  constructor(props) {
    super(props);

    this.state = {
      events: []
    };
  }

  async componentDidMount() {
    try {
      // Get the user's access token
      var accessToken = await (window.msal as any).acquireTokenSilent({
        scopes: config.scopes
      });
      // Get the user's events
      var events = await getEvents(accessToken);
      // Update the array of events in state
      this.setState({events: events.value});
    }
    catch(err) {
      console.log(err);
    }
  }

  render() {
    return (
      <div>
        <h1>Calendar</h1>
        <Table>
          <thead>
            <tr>
              <th scope="col">Organizer</th>
              <th scope="col">Subject</th>
              <th scope="col">Start</th>
              <th scope="col">End</th>
            </tr>
          </thead>
          <tbody>
            {this.state.events.map(
              function(event){
                return(
                  <tr key={event.id}>
                    <td>{event.organizer.emailAddress.name}</td>
                    <td>{event.subject}</td>
                    <td>{event.start.dateTime}</td>
                    <td>{event.end.dateTime}</td>
                  </tr>
                );
              })}
          </tbody>
        </Table>
      </div>
    );
  }
}