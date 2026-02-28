import {
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodeListSearchResult,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	NodeOperationError,
} from 'n8n-workflow';

/**
 * 自動轉換欄位值的型態
 * 支援：數字、布林值、日期、null
 */
function convertValueType(value: any): any {
	// 如果已經是 null、undefined，直接返回
	if (value === null || value === undefined) {
		return null;
	}

	// 如果不是字串，保持原樣
	if (typeof value !== 'string') {
		return value;
	}

	// 處理空字串
	if (value.trim() === '') {
		return null;
	}

	// 處理 "null" 字串
	if (value.toLowerCase() === 'null') {
		return null;
	}

	// 處理布林值
	if (value.toLowerCase() === 'true') {
		return true;
	}
	if (value.toLowerCase() === 'false') {
		return false;
	}

	// 處理數字（整數和浮點數）
	// 使用正則表達式確保是有效的數字格式
	if (/^-?\d+(\.\d+)?$/.test(value.trim())) {
		const num = Number(value);
		if (!isNaN(num)) {
			return num;
		}
	}

	// 處理日期格式
	// 只有日期（yyyy-MM-dd），直接回傳字串，不附加時間
	if (/^\d{4}-\d{2}-\d{2}$/.test(value.trim())) {
		return value.trim();
	}

	// 有時間的 ISO 格式（yyyy-MM-ddTHH:mm:ss...），才轉換為 ISO 字串
	if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(value.trim())) {
		const date = new Date(value);
		if (!isNaN(date.getTime())) {
			return date.toISOString();
		}
	}

	// 其他情況保持原字串
	return value;
}

/**
 * 遞迴轉換物件或陣列中的所有值
 */
function convertObjectValues(obj: any): any {
	if (obj === null || obj === undefined) {
		return null;
	}

	if (Array.isArray(obj)) {
		return obj.map(item => convertObjectValues(item));
	}

	if (typeof obj === 'object') {
		const converted: any = {};
		for (const [key, value] of Object.entries(obj)) {
			converted[key] = convertObjectValues(value);
		}
		return converted;
	}

	return convertValueType(obj);
}

export class ExcelApi implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Excel API',
		name: 'excelApi',
		icon: 'file:excelapi.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"]}}',
		description: 'Access Excel files via API with concurrent safety',
		defaults: {
			name: 'Excel API',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'excelApiAuth',
				required: true,
			},
		],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				options: [
					{ name: 'Append', value: 'append', action: 'Append row to Excel file', description: 'Add a new row to the end of the sheet' },
					{ name: 'Read', value: 'read', action: 'Read Excel file', description: 'Read data from Excel file' },
					{ name: 'Update', value: 'update', action: 'Update row', description: 'Update an existing row' },
					{ name: 'Delete', value: 'delete', action: 'Delete row', description: 'Delete a row' },
				],
				default: 'append',
			},
			// File selection
			{
				displayName: 'File Name',
				name: 'fileName',
				type: 'resourceLocator',
				required: true,
				default: { mode: 'list', value: '' },
				description: 'Select an Excel file from the server',
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						typeOptions: {
							searchListMethod: 'searchExcelFiles',
							searchable: true,
						},
					},
					{
						displayName: 'By Name',
						name: 'name',
						type: 'string',
						placeholder: 'e.g. report.xlsx',
						hint: 'Enter the exact Excel filename',
					},
				],
			},
			// Sheet selection
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'resourceLocator',
				default: { mode: 'list', value: '' },
				description: 'Select a worksheet from the file',
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						typeOptions: {
							searchListMethod: 'searchExcelSheets',
							searchable: true,
						},
					},
					{
						displayName: 'By Name',
						name: 'name',
						type: 'string',
						placeholder: 'e.g. Sheet1',
						hint: 'Enter the exact worksheet name',
					},
				],
			},
			// Append operation - Mode selection
			{
				displayName: 'Append Mode',
				name: 'appendMode',
				type: 'options',
				displayOptions: { 
					show: { 
						operation: ['append'] 
					} 
				},
				options: [
					{ name: 'Object (By Column Names)', value: 'object', description: 'Map values by column names - easier and safer' },
					{ name: 'Array (By Position)', value: 'array', description: 'Specify values in exact column order' },
				],
				default: 'object',
				description: 'How to specify values to append',
			},
			// Append - Object Mode
			{
				displayName: 'Values to Append',
				name: 'appendValuesObject',
				type: 'json',
				displayOptions: { 
					show: { 
						operation: ['append'],
						appendMode: ['object']
					} 
				},
				default: '{{ JSON.stringify({\n  "Column1": $json["field1"],\n  "Column2": $json["field2"]\n}) }}',
				required: true,
				description: 'Object with column names as keys. Column names must match Excel headers exactly.',
				hint: 'Example: {{ JSON.stringify({ "員工編號": $json["employeeId"], "姓名": $json["name"] }) }}',
			},
			// Append - Array Mode
			{
				displayName: 'Values to Append',
				name: 'appendValuesArray',
				type: 'json',
				displayOptions: { 
					show: { 
						operation: ['append'],
						appendMode: ['array']
					} 
				},
				default: '["value1", "value2", "value3"]',
				required: true,
				description: 'Array of values to append. Can use expressions like {{ $json.fieldName }}',
				hint: 'Example: ["{{ $json.name }}", "{{ $json.email }}", "{{ $json.age }}"]',
			},
			// Read operation
			{
				displayName: 'Range',
				name: 'range',
				type: 'string',
				displayOptions: { show: { operation: ['read'] } },
				default: '',
				description: 'Cell range to read (e.g., A1:D10). Leave empty to read all data',
				placeholder: 'A1:D10',
			},
			// Lookup Column
			{
				displayName: 'Lookup Column',
				name: 'lookupColumn',
				type: 'resourceLocator',
				default: { mode: 'list', value: '' },
				displayOptions: { 
					show: { 
						operation: ['update', 'delete'],
					} 
				},
				required: true,
				description: 'Column name to search in (automatically loaded from Excel headers)',
				hint: '💡 The list shows all column names from the first row of your Excel file',
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						typeOptions: {
							searchListMethod: 'searchColumnNames',
							searchable: true,
						},
					},
					{
						displayName: 'By Name',
						name: 'name',
						type: 'string',
						placeholder: 'e.g. EmployeeID',
						hint: 'Enter the exact column name from the first row of your Excel file',
					},
				],
			},
			// Lookup Value (for lookup method)
			{
				displayName: 'Lookup Value',
				name: 'lookupValue',
				type: 'string',
				displayOptions: { 
					show: { 
						operation: ['update', 'delete'],
					} 
				},
				required: true,
				default: '',
				placeholder: 'e.g., E001, john@example.com',
				description: 'Value to search for in the lookup column',
				hint: 'Can use expressions like {{ $json.id }}',
			},
			// Process Mode (for lookup method)
			{
				displayName: 'Process Mode',
				name: 'processMode',
				type: 'options',
				displayOptions: { 
					show: { 
						operation: ['update', 'delete'],
					} 
				},
				options: [
					{ 
						name: 'All Matching Records', 
						value: 'all', 
						description: 'Process all records that match the lookup condition' 
					},
					{ 
						name: 'First Match Only', 
						value: 'first', 
						description: 'Process only the first matching record' 
					},
				],
				default: 'all',
				description: 'Choose whether to process all matching records or just the first one',
				hint: '💡 "All Matching Records" will update/delete every row where the lookup column matches the lookup value',
			},
			// Update operation: Values to Set
			{
				displayName: 'Values to Set',
				name: 'valuesToSet',
				type: 'json',
				displayOptions: { show: { operation: ['update'] } },
				default: '{\n  "Status": "Done",\n  "UpdatedDate": "2024-01-01"\n}',
				required: true,
				description: 'Object with column names as keys and new values',
				hint: 'Example: {{ JSON.stringify({ "Status": $json["status"], "Salary": $json["salary"] }) }}'
			},
		],
	};

	methods = {
		loadOptions: {
			async getExcelSheets(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const fileNameRaw = this.getNodeParameter('fileName') as any;
				const fileName = (fileNameRaw && typeof fileNameRaw === 'object' ? fileNameRaw.value : fileNameRaw) as string;
				
				if (!fileName) {
					return [];
				}

				const credentials = await this.getCredentials('excelApiAuth');
				const apiUrl = credentials.url as string;
				const apiToken = credentials.token as string;

				try {
					const response = await this.helpers.request({
						method: 'GET',
						url: `${apiUrl}/api/excel/sheets?file=${encodeURIComponent(fileName)}`,
						headers: {
							'Authorization': `Bearer ${apiToken}`,
						},
						json: true,
					});

					if (response.success && response.sheets) {
						return response.sheets.map((sheet: string) => ({
							name: sheet,
							value: sheet,
						}));
					}
					return [];
				} catch (error) {
					return [];
				}
			},

			// ✨ NEW: 獲取欄位名稱（表頭）
			async getColumnNames(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const fileNameRaw = this.getNodeParameter('fileName') as any;
				const fileName = (fileNameRaw && typeof fileNameRaw === 'object' ? fileNameRaw.value : fileNameRaw) as string;
				const sheetNameRaw = this.getNodeParameter('sheetName') as any;
				const sheetName = (sheetNameRaw && typeof sheetNameRaw === 'object' ? sheetNameRaw.value : sheetNameRaw) as string;
				
				if (!fileName || !sheetName) {
					return [];
				}

				const credentials = await this.getCredentials('excelApiAuth');
				const apiUrl = credentials.url as string;
				const apiToken = credentials.token as string;

				try {
					const response = await this.helpers.request({
						method: 'GET',
						url: `${apiUrl}/api/excel/headers?file=${encodeURIComponent(fileName)}&sheet=${encodeURIComponent(sheetName)}`,
						headers: {
							'Authorization': `Bearer ${apiToken}`,
						},
						json: true,
					});

					if (response.success && response.headers) {
						return response.headers.map((header: string) => ({
							name: header,
							value: header,
						}));
					}
					return [];
				} catch (error) {
					// 如果獲取失敗，返回空列表（用戶可以手動輸入）
					return [];
				}
			},
		},
		listSearch: {
			async searchExcelFiles(this: ILoadOptionsFunctions, filter?: string): Promise<INodeListSearchResult> {
				const credentials = await this.getCredentials('excelApiAuth');
				const apiUrl = credentials.url as string;
				const apiToken = credentials.token as string;

				try {
					const response = await this.helpers.request({
						method: 'GET',
						url: `${apiUrl}/api/excel/files`,
						headers: { 'Authorization': `Bearer ${apiToken}` },
						json: true,
					});

					if (response.success && response.files) {
						let files: string[] = response.files;
						if (filter) {
							const f = filter.toLowerCase();
							files = files.filter((file: string) => file.toLowerCase().includes(f));
						}
						return {
							results: files.map((file: string) => ({ name: file, value: file })),
						};
					}
					return { results: [] };
				} catch (error) {
					return { results: [] };
				}
			},

			async searchExcelSheets(this: ILoadOptionsFunctions, filter?: string): Promise<INodeListSearchResult> {
				const fileNameRaw = this.getNodeParameter('fileName') as any;
				const fileName = (fileNameRaw && typeof fileNameRaw === 'object' ? fileNameRaw.value : fileNameRaw) as string;

				if (!fileName) {
					return { results: [] };
				}

				const credentials = await this.getCredentials('excelApiAuth');
				const apiUrl = credentials.url as string;
				const apiToken = credentials.token as string;

				try {
					const response = await this.helpers.request({
						method: 'GET',
						url: `${apiUrl}/api/excel/sheets?file=${encodeURIComponent(fileName)}`,
						headers: { 'Authorization': `Bearer ${apiToken}` },
						json: true,
					});

					if (response.success && response.sheets) {
						let sheets: string[] = response.sheets;
						if (filter) {
							const f = filter.toLowerCase();
							sheets = sheets.filter((s: string) => s.toLowerCase().includes(f));
						}
						return {
							results: sheets.map((sheet: string) => ({ name: sheet, value: sheet })),
						};
					}
					return { results: [] };
				} catch (error) {
					return { results: [] };
				}
			},

			async searchColumnNames(this: ILoadOptionsFunctions, filter?: string): Promise<INodeListSearchResult> {
				const fileNameRaw = this.getNodeParameter('fileName') as any;
				const fileName = (fileNameRaw && typeof fileNameRaw === 'object' ? fileNameRaw.value : fileNameRaw) as string;
				const sheetNameRaw = this.getNodeParameter('sheetName') as any;
				const sheetName = (sheetNameRaw && typeof sheetNameRaw === 'object' ? sheetNameRaw.value : sheetNameRaw) as string;

				if (!fileName || !sheetName) {
					return { results: [] };
				}

				const credentials = await this.getCredentials('excelApiAuth');
				const apiUrl = credentials.url as string;
				const apiToken = credentials.token as string;

				try {
					const response = await this.helpers.request({
						method: 'GET',
						url: `${apiUrl}/api/excel/headers?file=${encodeURIComponent(fileName)}&sheet=${encodeURIComponent(sheetName)}`,
						headers: { 'Authorization': `Bearer ${apiToken}` },
						json: true,
					});

					if (response.success && response.headers) {
						let headers: string[] = response.headers;
						if (filter) {
							const f = filter.toLowerCase();
							headers = headers.filter((h: string) => h.toLowerCase().includes(f));
						}
						return {
							results: headers.map((header: string) => ({ name: header, value: header })),
						};
					}
					return { results: [] };
				} catch (error) {
					return { results: [] };
				}
			},
		},
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];
		const operation = this.getNodeParameter('operation', 0) as string;

		// Get credentials
		const credentials = await this.getCredentials('excelApiAuth');
		const apiUrl = credentials.url as string;
		const apiToken = credentials.token as string;

		try {
			for (let i = 0; i < items.length; i++) {
				let responseData: any;

				// Common parameters (per-item to support Expression mode)
				const fileNameRaw = this.getNodeParameter('fileName', i) as any;
				const fileName = ((fileNameRaw && typeof fileNameRaw === 'object' ? fileNameRaw.value : fileNameRaw) as string) || '';
				const sheetNameRaw = this.getNodeParameter('sheetName', i) as any;
				const sheetName = ((sheetNameRaw && typeof sheetNameRaw === 'object' ? sheetNameRaw.value : sheetNameRaw) as string) || '';

				if (!fileName) {
					throw new NodeOperationError(
						this.getNode(),
						'File Name is required. Please select an Excel file.',
					);
				}

				if (operation === 'append') {
					// Get append mode to determine which parameter to use
					const appendMode = this.getNodeParameter('appendMode', i) as string;
					const parameterName = appendMode === 'object' ? 'appendValuesObject' : 'appendValuesArray';
					const appendValuesRaw = this.getNodeParameter(parameterName, i);
					let appendValues: any;

					if (typeof appendValuesRaw === 'string') {
						try {
							appendValues = JSON.parse(appendValuesRaw);
						} catch {
							throw new NodeOperationError(
								this.getNode(),
								`Values to Append must be a valid JSON ${appendMode === 'object' ? 'object' : 'array'}`,
							);
						}
						} else {
						appendValues = appendValuesRaw;
					}

					// 自動轉換值的型態
					const convertedValues = convertObjectValues(appendValues);

					// Use different API endpoint based on append mode
					if (appendMode === 'object') {
						// Object Mode: Use append_object API
						responseData = await this.helpers.request({
							method: 'POST',
							url: `${apiUrl}/api/excel/append_object`,
							headers: {
								'Authorization': `Bearer ${apiToken}`,
								'Content-Type': 'application/json',
							},
							body: {
								file: fileName,
								sheet: sheetName,
								values: convertedValues,
							},
							json: true,
						});
					} else {
						// Array Mode: Use standard append API
						responseData = await this.helpers.request({
							method: 'POST',
							url: `${apiUrl}/api/excel/append`,
							headers: {
								'Authorization': `Bearer ${apiToken}`,
								'Content-Type': 'application/json',
							},
							body: {
								file: fileName,
								sheet: sheetName,
								values: convertedValues,
							},
							json: true,
						});
					}
				} else if (operation === 'read') {
					const range = this.getNodeParameter('range', i) as string;

					responseData = await this.helpers.request({
						method: 'POST',
						url: `${apiUrl}/api/excel/read`,
						headers: {
							'Authorization': `Bearer ${apiToken}`,
							'Content-Type': 'application/json',
						},
						body: {
							file: fileName,
							sheet: sheetName,
							range: range || undefined,
						},
						json: true,
					});

					if (responseData.success && responseData.data) {
						const data = responseData.data as any[][];
						
						if (data.length <= 1) {
							returnData.push({ json: { success: true, data: [] } });
							continue;
						}

						if (data.length > 1) {
							const headers = data[0];
							const hasHeaders = headers.every((h: any) => typeof h === 'string' && h.length > 0);
							
							if (hasHeaders) {
								for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
									const rowData: any = {};
									const row = data[rowIdx];
									headers.forEach((header: string, colIdx: number) => {
										rowData[header] = row[colIdx];
									});
									returnData.push({ json: rowData });
								}
								continue;
							}
						}
						
						returnData.push({ json: responseData });
						continue;
					}

				} else if (operation === 'update') {
					const valuesToSetRaw = this.getNodeParameter('valuesToSet', i) as string;
					
					let valuesToSet: any;
					if (typeof valuesToSetRaw === 'string') {
						try {
							valuesToSet = JSON.parse(valuesToSetRaw);
						} catch {
							throw new NodeOperationError(
								this.getNode(),
								'Values to Set must be a valid JSON object',
							);
						}
					} else {
						valuesToSet = valuesToSetRaw;
					}

					// 自動轉換値的型態
					const convertedValuesToSet = convertObjectValues(valuesToSet);

					// Build request body
					const requestBody: any = {
						file: fileName,
						sheet: sheetName,
						values_to_set: convertedValuesToSet,
					};

					const lookupColumnRaw = this.getNodeParameter('lookupColumn', i) as any;
					const lookupColumn = (lookupColumnRaw && typeof lookupColumnRaw === 'object' ? lookupColumnRaw.value : lookupColumnRaw) as string;
					const lookupValue = this.getNodeParameter('lookupValue', i) as string;
					const processMode = this.getNodeParameter('processMode', i) as string;
					
					requestBody.lookup_column = lookupColumn;
					requestBody.lookup_value = lookupValue;
					requestBody.process_all = (processMode === 'all');

					responseData = await this.helpers.request({
						method: 'PUT',
						url: `${apiUrl}/api/excel/update_advanced`,
						headers: {
							'Authorization': `Bearer ${apiToken}`,
							'Content-Type': 'application/json',
						},
						body: requestBody,
						json: true,
					});

					// Check if any rows were affected
					if (responseData.success && responseData.updated_count === 0) {
						throw new NodeOperationError(
							this.getNode(),
							`No matching rows found. Lookup column: "${requestBody.lookup_column}", Lookup value: "${requestBody.lookup_value}"`,
						);
					}

				} else if (operation === 'delete') {
					// Build request body
					const requestBody: any = {
						file: fileName,
						sheet: sheetName,
					};

					const lookupColumnRaw = this.getNodeParameter('lookupColumn', i) as any;
					const lookupColumn = (lookupColumnRaw && typeof lookupColumnRaw === 'object' ? lookupColumnRaw.value : lookupColumnRaw) as string;
					const lookupValue = this.getNodeParameter('lookupValue', i) as string;
					const processMode = this.getNodeParameter('processMode', i) as string;
					
					requestBody.lookup_column = lookupColumn;
					requestBody.lookup_value = lookupValue;
					requestBody.process_all = (processMode === 'all');

					responseData = await this.helpers.request({
						method: 'DELETE',
						url: `${apiUrl}/api/excel/delete_advanced`,
						headers: {
							'Authorization': `Bearer ${apiToken}`,
							'Content-Type': 'application/json',
						},
						body: requestBody,
						json: true,
					});

					// Check if any rows were affected
					if (responseData.success && responseData.deleted_count === 0) {
						throw new NodeOperationError(
							this.getNode(),
							`No matching rows found. Lookup column: "${requestBody.lookup_column}", Lookup value: "${requestBody.lookup_value}"`,
						);
					}

				} else {
					throw new NodeOperationError(
						this.getNode(),
						`Unsupported operation: ${operation}`,
					);
				}

				returnData.push({ json: responseData });
			}
		} catch (error: any) {
			if (error.response?.body) {
				throw new NodeOperationError(
					this.getNode(),
					`Excel API Error: ${JSON.stringify(error.response.body)}`,
				);
			}
			throw new NodeOperationError(this.getNode(), error.message);
		}

		return [returnData];
	}
}
