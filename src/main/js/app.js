"use strict";

// tag::vars[]
import React from 'react';
import ExcelOperationPage from './component/ExcelOperationPage';

const ReactDOM = require("react-dom"); // <2>
const client = require("./client"); // <3>
import {
	Form,
	Button,
	message,
	notification,
	Divider,
	Row,
	Col,
	Upload,
	Modal,
} from "antd";
import { UploadOutlined } from "@ant-design/icons";
import axios from "axios";
import { Audio } from 'react-loader-spinner';
// end::vars[]

const FormItem = Form.Item;
// Create an Axios instance with the base URL
let apiContextPath = '';
if (process.env.NODE_ENV === 'production') {
	console.log('Running in production mode');
	console.log('API_CONTEXT value:', API_CONTEXT);
	apiContextPath = API_CONTEXT;
} else {
	console.log('Running in development mode');
	console.log('API_CONTEXT value:', '');
	apiContextPath = '';
}

const api = axios.create({
	baseURL: `${apiContextPath ? '/' + apiContextPath : ''}/api/v1`, // Set the base URL to match the configured base path in Spring
});

// tag::app[]
class App extends React.Component {
	// <1>
	constructor(props) {
		super(props);
		this.state = {};
	}

	componentDidMount() {
		// <2>
		/*client({method: 'GET', path: '/api/employees'}).done(response => {
				this.setState({employees: response.entity._embedded.employees});
			});*/
	}

	render() {
		// <3>

		return (
			<ExcelOperationPage />
		);
	}
}
// end::app[]

// tag::render[]
ReactDOM.render(<App />, document.getElementById("react"));
// end::render[]
