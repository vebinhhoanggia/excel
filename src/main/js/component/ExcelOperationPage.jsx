import React from "react";
import PropTypes from "prop-types";
import {
	Form, Button, message, notification, Divider,
	Row, Col, Upload
} from "antd";
import { UploadOutlined } from '@ant-design/icons';
import axios from "axios";
import { PulseLoader } from "halogenium";
import _ from 'lodash';

import { conf } from "../config/conf";
import authHeader from '../auth/auth-common';

const FormItem = Form.Item;

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

class ExcelOperationPage extends React.Component {
	formRef = React.createRef();
	static get propTypes() {
		return {
			form: PropTypes.object,
		};
	}

	constructor(props) {
		super(props);

		this.handleSubmit = this.handleSubmit.bind(this);

		this.state = {
			isLoading: false,
		};
	}

	componentDidMount() {
		document.title = 'HELLO';
	}

	onFinish = (values) => {
	};

	handleSubmit(e) {
		this.formRef.current.validateFields().then((fieldsValue) => {
			// const url = `${conf.operation.excel.split}`;
			const url = '/opeation/excel/upload-excel';

			const { file1List: file1 } = this.state;
			const formData = new FormData();

			for (let i = 0; i < file1.length; i++) {
				if (file1[i]) {
					formData.append('files', file1[i]);
					// formData.append('files', file1[i], encodeURIComponent(file1[i].name));
				}
			}


			this.setState(
				{
					isLoading: true,
				},
				() => { }
			);
			api
				.post(url, formData, {
					headers: {
						'Content-Type': 'multipart/form-data',
					},
					responseType: 'blob'
				})
				.then((response) => {
					const config = {
						title: 'title',
						content: '',
						okText: 'ＯＫ',
						onOk() {
							() => { };
						},
					};
					this.openModalInfo(config);

					// const blob = new Blob([response.data], { type: response.headers['content-type'] });
					const filename = this.extractFileName(
						response.headers["content-disposition"]
					);
					const downloadUrl = window.URL.createObjectURL(new Blob([response.data]));
					// const downloadUrl = URL.createObjectURL(blob);
					const link = document.createElement('a');
					link.href = downloadUrl;
					link.setAttribute('download', filename);
					document.body.appendChild(link);
					link.click();
					link.remove();
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
		const { file1List, isLoading } = this.state;
		const className = isLoading ? "loading" : "";
		const overlay = isLoading ? "overlay" : "";
		const defaultProps = {
			// maxCount: 1,
			multiple: true,
		};
		const upload1Props = {
			...defaultProps,
			onRemove: (f) => this.handleRemove(1, f),
			beforeUpload: (file) => {
				this.setState((prevState) => ({
					file1List: [...prevState.file1List, file],
				}));
				return false;
			},
			file1List,
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
				<h1>Tach file excel</h1>
				<Form
					layout="horizontal"
					initialValues={{
					}}
					ref={this.formRef}
					name="control-ref"
					onFinish={this.onFinish}
					labelCol={{ span: 8 }}
					wrapperCol={{ span: 16 }}
					className='qms-form'
				>
					<FormItem wrapperCol={{ span: 24 }}>
						<div className="export-title">Ố LA LA</div>
					</FormItem>
					<Form.Item
						name="file1"
						label="INPUT"
						valuePropName="file1List"
						getValueFromEvent={(e) => e.file1List}
					>
						<Upload {...upload1Props}>
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
									SplitFile
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

export default ExcelOperationPage;
