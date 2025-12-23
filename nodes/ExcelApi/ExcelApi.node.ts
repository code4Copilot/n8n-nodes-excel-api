import {
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	NodeOperationError,
} from 'n8n-workflow';

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
					{ name: 'Batch', value: 'batch', action: 'Batch operations', description: 'Execute multiple operations at once' },
				],
				default: 'append',
			},
			// File selection
			{
				displayName: 'File Name',
				name: 'fileName',
				type: 'options',
				required: true,
				typeOptions: {
					loadOptionsMethod: 'getExcelFiles',
				},
				default: '',
				description: 'Select an Excel file from the server',
			},
			// Sheet selection
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getExcelSheets',
					loadOptionsDependsOn: ['fileName'],
				},
				default: 'Sheet1',
				description: 'Select a worksheet from the file',
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
				default: '{\n  "Column1": "{{ $json.field1 }}",\n  "Column2": "{{ $json.field2 }}"\n}',
				required: true,
				description: 'Object with column names as keys. Column names must match Excel headers exactly.',
				hint: 'Example: {"員工編號": "{{ $json.body.employeeId }}", "姓名": "{{ $json.body.name }}"}',
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
			// Update & Delete: Row Identification Method
			{
				displayName: 'Identify Row By',
				name: 'identifyBy',
				type: 'options',
				displayOptions: { show: { operation: ['update', 'delete'] } },
				options: [
					{ name: 'Row Number', value: 'rowNumber', description: 'Specify the exact row number' },
					{ name: 'Lookup', value: 'lookup', description: 'Find row by matching a column value' },
				],
				default: 'rowNumber',
				description: 'How to identify the row to update/delete',
			},
			// Row Number (for direct specification)
			{
				displayName: 'Row Number',
				name: 'rowNumber',
				type: 'number',
				displayOptions: { 
					show: { 
						operation: ['update', 'delete'],
						identifyBy: ['rowNumber']
					} 
				},
				required: true,
				default: 2,
				description: 'Row number to update/delete (1-based, row 1 is header)',
				hint: '⚠️ Row 1 is protected (header row). Data rows start from row 2.',
			},
			// Lookup Column (for lookup method)
			{
				displayName: 'Lookup Column',
				name: 'lookupColumn',
				type: 'string',
				displayOptions: { 
					show: { 
						operation: ['update', 'delete'],
						identifyBy: ['lookup']
					} 
				},
				required: true,
				default: '',
				placeholder: 'e.g., 員工編號, Email, ID',
				description: 'Column name to search in (must match header exactly)',
				hint: 'The first row is treated as headers',
			},
			// Lookup Value (for lookup method)
			{
				displayName: 'Lookup Value',
				name: 'lookupValue',
				type: 'string',
				displayOptions: { 
					show: { 
						operation: ['update', 'delete'],
						identifyBy: ['lookup']
					} 
				},
				required: true,
				default: '',
				placeholder: 'e.g., E001, john@example.com',
				description: 'Value to search for in the lookup column',
				hint: 'Can use expressions like {{ $json.id }}',
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
				hint: 'Example: {"Status": "{{ $json.status }}", "Salary": {{ $json.salary }}}',
			},
			// Batch operation
			{
				displayName: 'Operations',
				name: 'batchOperations',
				type: 'json',
				displayOptions: { show: { operation: ['batch'] } },
				default: `[
  {
    "type": "append",
    "values": ["value1", "value2"]
  },
  {
    "type": "update",
    "row": 5,
    "values": ["new1", "new2"]
  }
]`,
				required: true,
				description: 'Array of operations to execute',
				hint: 'Each operation should have "type" (append/update/delete) and related fields',
			},
		],
	};

	methods = {
		loadOptions: {
			async getExcelFiles(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const credentials = await this.getCredentials('excelApiAuth');
				const apiUrl = credentials.url as string;
				const apiToken = credentials.token as string;

				try {
					const response = await this.helpers.request({
						method: 'GET',
						url: `${apiUrl}/api/excel/files`,
						headers: {
							'Authorization': `Bearer ${apiToken}`,
						},
						json: true,
					});

					if (response.success && response.files) {
						return response.files.map((file: string) => ({
							name: file,
							value: file,
						}));
					}
					return [];
				} catch (error) {
					return [];
				}
			},

			async getExcelSheets(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const fileName = this.getNodeParameter('fileName') as string;
				
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

		// Common parameters
		const fileName = this.getNodeParameter('fileName', 0) as string;
		const sheetName = this.getNodeParameter('sheetName', 0) as string || 'Sheet1';

		if (!fileName) {
			throw new NodeOperationError(
				this.getNode(),
				'File Name is required. Please select an Excel file.',
			);
		}

		try {
			for (let i = 0; i < items.length; i++) {
				let responseData: any;

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

					// Use different API endpoint based on append mode
					if (appendMode === 'object') {
						// Object Mode: Use append_object API
						// This API automatically reads headers and maps values by column names
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
								values: appendValues,
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
								values: appendValues,
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
					// Get identification method
					const identifyBy = this.getNodeParameter('identifyBy', i) as string;
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

					// Build request body
					const requestBody: any = {
						file: fileName,
						sheet: sheetName,
						values_to_set: valuesToSet,
					};

					if (identifyBy === 'rowNumber') {
						const rowNumber = this.getNodeParameter('rowNumber', i) as number;
						requestBody.row = rowNumber;
					} else if (identifyBy === 'lookup') {
						const lookupColumn = this.getNodeParameter('lookupColumn', i) as string;
						const lookupValue = this.getNodeParameter('lookupValue', i) as string;
						requestBody.lookup_column = lookupColumn;
						requestBody.lookup_value = lookupValue;
					}

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

				} else if (operation === 'delete') {
					// Get identification method
					const identifyBy = this.getNodeParameter('identifyBy', i) as string;

					// Build request body
					const requestBody: any = {
						file: fileName,
						sheet: sheetName,
					};

					if (identifyBy === 'rowNumber') {
						const rowNumber = this.getNodeParameter('rowNumber', i) as number;
						requestBody.row = rowNumber;
					} else if (identifyBy === 'lookup') {
						const lookupColumn = this.getNodeParameter('lookupColumn', i) as string;
						const lookupValue = this.getNodeParameter('lookupValue', i) as string;
						requestBody.lookup_column = lookupColumn;
						requestBody.lookup_value = lookupValue;
					}

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

				} else if (operation === 'batch') {
					const batchOperationsRaw = this.getNodeParameter('batchOperations', i) as string;
					
					let batchOperations: any[];
					if (typeof batchOperationsRaw === 'string') {
						try {
							batchOperations = JSON.parse(batchOperationsRaw);
						} catch {
							throw new NodeOperationError(
								this.getNode(),
								'Batch Operations must be a valid JSON array',
							);
						}
					} else {
						batchOperations = batchOperationsRaw;
					}

					responseData = await this.helpers.request({
						method: 'POST',
						url: `${apiUrl}/api/excel/batch`,
						headers: {
							'Authorization': `Bearer ${apiToken}`,
							'Content-Type': 'application/json',
						},
						body: {
							file: fileName,
							sheet: sheetName,
							operations: batchOperations,
						},
						json: true,
					});

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
