# Copyright (c) 2023 Marcin Bugała


"""liblary imports and class definition"""
try:
    import adal
except: 
    ! pip install adal
    import adal
import json, requests, pandas as pd
from datetime import datetime
import logging, time


class Workspace:
    """
    A class to represent a workspace in Power BI.
    ...
    Class contains Workspaces ID dictionary containing group_id for all main workspaces.
    ...
    To create instance of class, call it with workspace name, ex: Workspace('Sales Reports') or use group_ID.
    User 'Service Principal' must be admin in that workspace. 
    Default = 'Sales Reports'
    
    Attributes
    ----------
    name : str
        Name of workspace, default = 'Sales Reports'
    group_id : str
        The workspace ID
    header : str
        Authorization header with connection token
    datasets: dict
        Dictionary with all dasaets in workspace
        {name: Dataset} where Dataset is Dataset class object
        
    Methods
    -------
    get_token(self) : called in __init__
        Method used to setup connection with PowerBI Rest API, and create self.header
    get_datasets(self) : called in __init__
        Method used to get all datasets to create dictionary self.datasets
    get_all_workspaces(self) :
        Get dictionary with all available workspaces for Service Principal Azure App
        
    Subclass
    --------
    Dataset :
        Object representing dataset in workspace
    """
    # you can define your own dictionary with workspaces names & ID`s
    workspaces_id = {'Sales Reports': 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'}
    
    
    def __init__(self, name = 'Sales Reports'):
        """
        Constructs all the necessary attributes for the workspace object.
        
        Parameters
        ----------
        name or group_id : str
            Name or group_id of workspace
        """
        
        self.header = {'Authorization': f'Bearer {self.get_token()}'}
            
        if name in self.workspaces_id:
            self.group_id = self.workspaces_id[name]
            self.name = name
        else:
            self.group_id = name
            self.name = list(self.get_all_workspaces().keys())[list(self.get_all_workspaces().values()).index(name)]
            
        self.datasets = self.get_datasets()

            
    def __repr__(self):
        return f'Workspace: {self.name}'


    def get_token(self):
        """
        Method used to setup connection with PowerBI Rest API, and create self.header
        """

        # below provide your own tenant_id, client_id and client_secret
        tenant_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
        authority_url = 'https://login.microsoftonline.com/'+tenant_id+'/'
        resource_url = 'https://analysis.windows.net/powerbi/api'
        client_secret = "powerbirestAPIAppSecret" 
        client_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
        context = adal.AuthenticationContext(authority=authority_url,
                                         validate_authority=True,
                                         api_version=None)
        token = context.acquire_token_with_client_credentials(resource_url, client_id, client_secret)

        return token.get('accessToken')

    
    def get_datasets(self):
        """
        Method used to get all datasets to create dictionary self.datasets
        
        Rest API: GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets
        """

        result = {}
        url = 'https://api.powerbi.com/v1.0/myorg/groups/' + self.group_id + '/datasets'
        result_json = json.loads( requests.get(url, headers = self.header ).content)['value']

        for x in result_json:
            if x['name'] == 'Report Usage Metrics Model':
                pass
            else:
                result[x['name']] = self.Dataset(self.group_id ,x['name'], x['id'], self.header)

        return result
    
    
    def get_all_workspaces(self):
        """
        Get dictionary with all available workspaces for Service Principal Azure App
        
        Rest API: GET https://api.powerbi.com/v1.0/myorg/groups
        """
        
        result = {}
        url = 'https://api.powerbi.com/v1.0/myorg/groups/'
        result_json = json.loads( requests.get(url, headers = self.header ).content)['value']

        for x in result_json:
                result[x['name']] =  x['id']

        return result

    
    # subclass definition


    class Dataset:
        """
        A class to represent a dataset object in workspace object.
        ...
        
        Attributes
        ----------
        group_id : str
            The workspace ID
        name : str
            Name of dataset
        dataset_id : str
            The dataset ID
        header : str
            Authorization header with connection token
        refreshes : list
            List with dictionaries describing latest refresh
            [{'status': 'Completed', 'refreshType': 'ViaApi', 'duration': '0:02:06.726000', 'endTime': '2023-03-02 04:42:02.763000'}]

        Methods
        -------
        get_refreshes(top = 1) :
            Get list with dictionaries describing top latest refreshes. Default top = 1
        execute_query(query) : 
            Execute DAX query against dataset, ex:
            query = 'EVALUATE ROW("Sales Amount", \'Measures Table\'[Data Update Monitoring])'
        refresh(table = "")
            Refresh dataset. On Premium Capacity or PPU you may provide select table to refresh.
        """
    
    
        def __init__(self, group_id, name, dataset_id, header):
            """
            Constructs all the necessary attributes for the dataset object.
            
            Parameters
            ----------
            group_id : str
                The workspace ID
            name : str
                Name of dataset
            dataset_id : str
                The dataset ID
            header : str
                Authorization header with connection token
            
            """
            self.group_id = group_id
            self.name = name
            self.dataset_id = dataset_id         
            self.header = header
            self.refreshes = self.get_refreshes()
            
            
        def __repr__(self):
            return f'Datset: {self.name}'
        
        
        def get_refreshes(self, top = 1):
            """
            Get list with dictionaries describing top latest refreshes. Default top = 1
            
            Rest API: GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes?$top={$top}
            """ 

            result = []
            url = f'https://api.powerbi.com/v1.0/myorg/groups/{self.group_id}/datasets/{self.dataset_id}/refreshes?$top={top}'
            try:
                result_json = json.loads( requests.get(url, headers = self.header ).content)['value']
                for x in result_json:
                    try:
                        startTime = datetime.strptime(x['startTime'], '%Y-%m-%dT%H:%M:%SZ')
                    except ValueError:
                        startTime = datetime.strptime(x['startTime'], '%Y-%m-%dT%H:%M:%S.%fZ')

                    try:
                        endTime = datetime.strptime(x['endTime'], '%Y-%m-%dT%H:%M:%SZ')
                    except ValueError:                        
                        endTime = datetime.strptime(x['endTime'], '%Y-%m-%dT%H:%M:%S.%fZ')

                    duration = endTime - startTime
                    result.append({'status': x['status'], 'refreshType': x['refreshType'], 'duration': f'{duration}', 'endTime': f'{endTime}'})
            except KeyError:
                pass

            return result
        
        
        def execute_query(self, query):
            """
            Execute DAX query against dataset, ex:
            query = 'EVALUATE ROW("Sales Amount", \'Measures Table\'[Data Update Monitoring])'
            
            Rest API: POST https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/executeQueries
            """

            query_json = {"queries":[{"query": query}]}
            url = f'https://api.powerbi.com/v1.0/myorg/groups/{self.group_id}/datasets/{self.dataset_id}/executeQueries'
            response = requests.post( url, headers = self.header, json = query_json)
            response_json = json.loads( response.content )
            try:
                result_json = response_json['results'][0]['tables'][0]['rows'][0]
            except KeyError:
                result_json = f"{response.status_code}: {response_json['error']}"
            except:
                result_json = response.status_code

            return result_json
        
        
        def refresh(self, table = ""):
            """
            Refresh dataset. On Premium Capacity or PPU you may provide select table to refresh.
            
            Rest API: POST https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes
            """
            
            url = f'https://api.powerbi.com/v1.0/myorg/groups/{self.group_id}/datasets/{self.dataset_id}/refreshes'
            if table == "":
                response = requests.post( url, headers = self.header)
            else:
                refresh_json = {"type": "full", "objects": [{"table": table }], "applyRefreshPolicy": "false" }
                response = requests.post( url, headers = self.header, json = refresh_json)
                
            if response.status_code == 202:
                return '202 Accepted'
            else:
                code = response.status_code
                message = json.loads(response.content)['error']['message']
                return f'Code: {code}, {message}'

      
        def get_tables(self):
            """
            Get list of table names in dataset.

            Rest API: POST https://api.powerbi.com/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/executeQueries
            Query: EVALUATE SUMMARIZE (COLUMNSTATISTICS(), [Table Name], [Column Name])
            """

            query_json = {"queries":[{"query": 'EVALUATE SUMMARIZE (COLUMNSTATISTICS(), [Table Name], [Column Name])'}]}
            url = f'https://api.powerbi.com/v1.0/myorg/groups/{self.group_id}/datasets/{self.dataset_id}/executeQueries'
            response = requests.post( url, headers = self.header, json = query_json)

            return list(dict.fromkeys( [x['[Table Name]'] for x in json.loads(response.content)['results'][0]['tables'][0]['rows']] ))
    
    
    # end of subclass definition    
