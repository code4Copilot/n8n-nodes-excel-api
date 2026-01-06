import { IExecuteFunctions, ILoadOptionsFunctions } from 'n8n-workflow';

export class MockExecuteFunctions {
  private parameters: { [key: string]: any } = {};
  private credentials: { [key: string]: any } = {};
  private inputData: any[] = [];
  private nodeData: any = {};
  private requestResponses: Map<string, any> = new Map();
  private lastRequestBody: any = null;

  constructor() {
    this.nodeData = {
      name: 'Excel API Test',
      type: 'n8n-nodes-excel-api.excelApi',
    };
  }

  setParameter(name: string, value: any, itemIndex: number = 0): void {
    if (!this.parameters[itemIndex]) {
      this.parameters[itemIndex] = {};
    }
    this.parameters[itemIndex][name] = value;
  }

  setCredentials(name: string, credentials: any): void {
    this.credentials[name] = credentials;
  }

  setInputData(data: any[]): void {
    this.inputData = data;
  }

  setRequestResponse(url: string, response: any): void {
    this.requestResponses.set(url, response);
  }

  getLastRequestBody(): any {
    return this.lastRequestBody;
  }

  getExecuteFunctions(): IExecuteFunctions {
    return {
      getNodeParameter: (parameterName: string, itemIndex: number) => {
        return this.parameters[itemIndex]?.[parameterName];
      },
      getCredentials: async (type: string) => {
        return this.credentials[type];
      },
      getInputData: () => {
        return this.inputData;
      },
      getNode: () => {
        return this.nodeData;
      },
      helpers: {
        request: async (options: any) => {
          const response = this.requestResponses.get(options.url);
          if (!response) {
            throw new Error(`No mock response set for URL: ${options.url}`);
          }
          if (response.error) {
            throw response.error;
          }
          // Store the request body for testing
          this.lastRequestBody = options.body;
          return response;
        },
      },
    } as any;
  }

  getLoadOptionsFunctions(): ILoadOptionsFunctions {
    return {
      getCredentials: async (type: string) => {
        return this.credentials[type];
      },
      getNodeParameter: (parameterName: string) => {
        return this.parameters[0]?.[parameterName];
      },
      helpers: {
        request: async (options: any) => {
          const response = this.requestResponses.get(options.url);
          if (!response) {
            throw new Error(`No mock response set for URL: ${options.url}`);
          }
          if (response.error) {
            throw response.error;
          }
          return response;
        },
      },
    } as any;
  }
}
