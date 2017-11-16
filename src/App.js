import React, { Component } from 'react';
import './App.css';

class App extends Component {
  constructor(props) {
    super(props);

    this.onColorMe = this.onColorMe.bind(this);
  }

  onColorMe() {
    window.Excel.run((context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = 'green';
      return context.sync();
    });
  }

  render() {
    return (
      <div id="content">
        <div id="content-header">
          <div className="padding">
              <h1>Welcome</h1>
          </div>
        </div>
        <div id="content-main">
          <div className="padding">
              <p>Choose the button below to set the color of the selected range to green.</p>
              <br />
              <h3>Try it out</h3>
              <button onClick={this.onColorMe}>Color Me</button>
          </div>
        </div>
      </div>
    );
  }
}

export default App;
