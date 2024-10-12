param logicAppName string = 'WeeklyCalendarEventRetrieval'
param location string = resourceGroup().location
param office365ConnectionName string = 'office365'
param teamsConnectionName string = 'teams'

resource office365Connection 'Microsoft.Web/connections@2016-06-01' = {
  name: office365ConnectionName
  location: location
  properties: {
    displayName: 'Office 365'
    api: {
      id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'office365')
    }
  }
}

resource teamsConnection 'Microsoft.Web/connections@2016-06-01' = {
  name: teamsConnectionName
  location: location
  properties: {
    displayName: 'Microsoft Teams'
    api: {
      id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'teams')
    }
  }
}

resource logicApp 'Microsoft.Logic/workflows@2019-05-01' = {
  name: logicAppName
  location: location
  properties: {
    state: 'Enabled'
    definition: {
      '$schema': 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#'
      contentVersion: '1.0.0.0'
      parameters: {
        '$connections': {
          defaultValue: {}
          type: 'Object'
        }
      }
      triggers: {
        Recurrence: {
          recurrence: {
            frequency: 'Week'
            interval: 1
            schedule: {
              hours: [
                '15'
              ]
              minutes: [
                0
              ]
              weekDays: [
                'Friday'
              ]
            }
          }
          type: 'Recurrence'
        }
      }
      actions: {
        Get_calendar_events: {
          inputs: {
            host: {
              connection: {
                name: '@parameters(\'$connections\')[\'office365\'][\'connectionId\']'
              }
            }
            method: 'get'
            path: '/v2/me/calendarview'
            queries: {
              startDateTime: '@{formatDateTime(addDays(startOfWeek(utcNow(), \'Monday\'), -7), \'yyyy-MM-ddTHH:mm:ss\')}'
              endDateTime: '@{formatDateTime(utcNow(), \'yyyy-MM-ddTHH:mm:ss\')}'
            }
          }
          runAfter: {}
          type: 'ApiConnection'
        }
        Process_events: {
          type: 'Foreach'
          foreach: '@body(\'Get_calendar_events\')?[\'value\']'
          actions: {
            Calculate_duration: {
              type: 'Compose'
              inputs: '@div(sub(ticks(item()?[\'end\']?[\'dateTime\']), ticks(item()?[\'start\']?[\'dateTime\'])), 36000000000)'
            }
            Format_event_list: {
              type: 'Compose'
              inputs: '@join(map(groupBy(body(\'Create_event_list\'), \'date\'), (group) => concat(formatDateTime(parse(group[\'key\']), \'dddd\'), \', \', group[\'key\'], \'\n\', join(map(group[\'value\'], (event) => concat(\'- \', event[\'title\'], \' (Duration: \', string(event[\'duration\']), \' hours, Category: \', event[\'category\'], \')\n\')), \'\'))), \'\n\n\')'
              runAfter: {
                Create_event_list: [
                  'Succeeded'
                ]
              }
            }          }
          runAfter: {
            Get_calendar_events: [
              'Succeeded'
            ]
          }
        }
        Create_event_list: {
          type: 'Select'
          inputs: {
            from: '@body(\'Get_calendar_events\')?[\'value\']'
            select: {
              date: '@formatDateTime(item()?[\'start\']?[\'dateTime\'], \'yyyy-MM-dd\')'
              title: '@item()?[\'subject\']'
              duration: '@outputs(\'Calculate_duration\')'
              category: '@outputs(\'Categorize_event\')'
            }
          }
          runAfter: {
            Process_events: [
              'Succeeded'
            ]
          }
        }
        Format_event_list: {
          type: 'Compose'
          inputs: '@join(map(groupBy(body(\'Create_event_list\'), \'date\'), (group) => concat(formatDateTime(parse(group[\'key\']), \'dddd\'), \', \', group[\'key\'], \'\n\', join(map(group[\'value\'], (event) => concat(\'- \', event[\'title\'], \' (Duration: \', string(event[\'duration\']), \' hours, Category: \', event[\'category\'], \')\n\')), \'\'))), \'\n\n\')'
          runAfter: {
            Create_event_list: [
              'Succeeded'
            ]
          }
        }
        Send_Teams_message: {
          type: 'ApiConnection'
          inputs: {
            host: {
              connection: {
                name: '@parameters(\'$connections\')[\'teams\'][\'connectionId\']'
              }
            }
            method: 'post'
            path: '/v3/beta/teams/conversation/message/poster/Flow bot/location/@{encodeURIComponent(\'User\')}'
            body: {
              messageBody: 'Weekly Calendar Event Summary:\n\n@{outputs(\'Format_event_list\')}'
            }
          }
          runAfter: {
            Format_event_list: [
              'Succeeded'
            ]
          }
        }
      }
      outputs: {}
    }
    parameters: {
      '$connections': {
        value: {
          office365: {
            connectionId: office365Connection.id
            connectionName: 'office365'
            id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'office365')
          }
          teams: {
            connectionId: teamsConnection.id
            connectionName: 'teams'
            id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'teams')
          }
        }
      }
    }
  }
}

output logicAppResourceId string = logicApp.id
