import {
	ICredentialType,
	INodeProperties,
} from 'n8n-workflow';

export class ExcelApiAuth implements ICredentialType {
	name = 'excelApiAuth';
	displayName = 'Excel API Auth';
	documentationUrl = 'https://your-docs-url.com';
	properties: INodeProperties[] = [
		{
			displayName: 'API URL',
			name: 'url',
			type: 'string',
			default: 'http://localhost:8000',
			required: true,
			description: 'The base URL of your Excel API server',
			placeholder: 'http://localhost:8000',
		},
		{
			displayName: 'API Token',
			name: 'token',
			type: 'string',
			typeOptions: {
				password: true,
			},
			default: '',
			required: true,
			description: 'The API token for authentication',
		},
	];
}