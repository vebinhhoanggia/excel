"use strict";

// tag::vars[]
import React from 'react';
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
const api = axios.create({
	baseURL: "/backlogApp/api/v1", // Set the base URL to match the configured base path in Spring
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
			<GenSchedule />
		);
	}
}
// end::app[]

// tag::GenSchedule[]
class GenSchedule extends React.Component {
	// <1>
	formRef = React.createRef();
	constructor(props) {
		super(props);
		this.handleSubmit = this.handleSubmit.bind(this);
		this.state = { file1List: [], file2List: [] };
	}

	modal = null;

	onFinish = (values) => { };

	openModalInfo = (config) => {
		Modal.info(config);
	};

	handleSubmit(e) {
		this.formRef.current.validateFields().then((fieldsValue) => {
			// Should format date value before submit.
			const url = '/base/genSchedule';

			const {
				file1List: file1,
				file2List: file2,
			} = this.state;
			const formData = new FormData();

			let file1Val = null;
			if (file1 && file1.length) {
				for (let i = 0; i < file1.length; i++) {
					if (file1[i]) {
						file1Val = file1[i];
						break;
					}
				}
			}
			formData.append("file1", file1Val);

			let file2Val = null;
			if (file2 && file2.length) {
				for (let i = 0; i < file2.length; i++) {
					if (file2[i]) {
						file2Val = file2[i];
						break;
					}
				}
			}
			formData.append("file2", file2Val);

			this.setState(
				{
					isLoading: true,
				},
				() => { }
			);
			api
				.post(url, formData, {
					headers: {
						"Content-Type": "multipart/form-data",
					},
				})
				.then((response) => {
					const config = {
						title: 'title',
						content: response.data,
						okText: 'ＯＫ',
						onOk() {
							() => { };
						},
					};
					this.openModalInfo(config);
				})
				.catch((error) => {
					// handle error
					message.error(error.message, [1], null);
					if (error.response) {
						notification.error({
							key: "init",
							message: "Error",
							description: error.response.data.message,
						});
						// The request was made and the server responded with a status code
						// that falls out of the range of 2xx
					} else if (error.request) {
						// The request was made but no response was received
						// `error.request` is an instance of XMLHttpRequest in the browser and an instance of
						// http.ClientRequest in node.js
					} else {
						// Something happened in setting up the request that triggered an Error
					}
				})
				.finally(() => {
					// always executed
					this.setState(
						{
							isLoading: false,
						},
						() => { }
					);
				});
		});
	}

	handleRemove(k, file) {
		const fieldName = `file${k}List`;
		const fileList = this.state[fieldName];
		const index = fileList.indexOf(file);
		const newFileList = fileList.slice();
		newFileList.splice(index, 1);
		this.setState({ [fieldName]: newFileList });
	}

	render() {
		// <3>
		const { file1List, file2List, isLoading } = this.state;
		const className = isLoading ? "loading" : "";
		const overlay = isLoading ? "overlay" : "";
		const defaultProps = {
			maxCount: 1,
		};
		const upload1Props = {
			...defaultProps,
			onRemove: (f) => this.handleRemove(1, f),
			beforeUpload: (file) => {
				const pattern = /^pjjyuji_data_csv_\d{8}\.csv$/;

				if (!pattern.test(file.name)) {
					message.error(
						'Invalid file name pattern. Please use "pjjyuji_data_csv_yyyymmdd.csv" format.'
					);
					return Upload.LIST_IGNORE;
				}
				this.setState((prevState) => ({
					file1List: [file],
				}));
				return false;
			},
			file1List,
		};
		const upload2Props = {
			...defaultProps,
			onRemove: (f) => this.handleRemove(2, f),
			beforeUpload: (file) => {
				const pattern = /^Backlog-Issues-\d{8}-\d{4}\.csv$/;

				if (!pattern.test(file.name)) {
					message.error(
						'Invalid file name pattern. Please use "Backlog-Issues-yyyymmdd-0000.csv" format.'
					);
					return Upload.LIST_IGNORE;
				}
				this.setState((prevState) => ({
					file2List: [file],
				}));
				return false;
			},
			file2List,
		};

		return (
			<div>
				{(() => {
					if (this.state.isLoading) {
						return (
							<div className={overlay}>
								<div
									ref={(node) => {
										this.loading = node;
									}}
									className={className}
									id="load"
								>
									{<Audio
										height="80"
										width="80"
										radius="9"
										color="green"
										ariaLabel="loading"
										wrapperStyle
										wrapperClass
									/>}
								</div>
							</div>
						);
					}
					return "";
				})()}
				<h1>BACKLOG STASTICS</h1>
				<Form
					layout="horizontal"
					initialValues={{}}
					ref={this.formRef}
					name="control-ref"
					onFinish={this.onFinish}
					labelCol={{ span: 8 }}
					wrapperCol={{ span: 16 }}
					className="qms-form"
				>
					<Form.Item
						name="file1"
						label="Pjjyuji"
						valuePropName="file1List"
						getValueFromEvent={(e) => e.file1List}
					>
						<Upload {...upload1Props}>
							<Button icon={<UploadOutlined />}>Select Files</Button>
						</Upload>
					</Form.Item>

					<Form.Item
						name="file2"
						label="Issues"
						valuePropName="file2List"
						getValueFromEvent={(e) => e.file2List}
					>
						<Upload {...upload2Props}>
							<Button icon={<UploadOutlined />}>Select Files</Button>
						</Upload>
					</Form.Item>

					<Row>
						<Col span={9} />
						<Col span={6}>
							<FormItem
								wrapperCol={{
									xs: { span: 24, offset: 0 },
									sm: { span: 16, offset: 8 },
								}}
							>
								<Button
									className="btn-export"
									type="primary"
									onClick={this.handleSubmit}
									size="large"
								>
									generate
								</Button>
							</FormItem>
						</Col>
					</Row>
				</Form>
				<Divider />
			</div>
		);
	}
}
// end::GenSchedule[]

// tag::employee-list[]

// tag::render[]
ReactDOM.render(<App />, document.getElementById("react"));
// end::render[]
