// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// Generated with Bot Builder V4 SDK Template for Visual Studio EchoBot v4.9.1
extern alias BetaLib;
using Beta = BetaLib.Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Bot.Builder.Community.Samples.Teams.Models;
using Bot.Builder.Community.Samples.Teams.Services;
using Microsoft.Bot.Connector;
using Newtonsoft.Json.Linq;
using Microsoft.Bot.Connector.Authentication;
using Newtonsoft.Json;
using System.IO;

using Microsoft.Graph;
using CustomHelperClasses;
using CommonModels;
using System.Text;
using System.Globalization;

namespace Bot.Builder.Community.Samples.Teams.Bots
{
    public class EchoBot : TeamsActivityHandler
    {
        readonly string _connectionName;
        private string plannerGroupId;
        public static string botClientID;
        public static string botClientSecret;
        private string tenantId;
        private string serviceUrl;
        private string tenantDomainName;
        private string stations;

        public EchoBot(IConfiguration configuration)
        {
            _connectionName = configuration["ConnectionNameGraph"] ?? throw new NullReferenceException("ConnectionNameGraph");
            botClientID = configuration["MicrosoftAppId"];
            botClientSecret = configuration["MicrosoftAppPassword"];
            tenantId = configuration["tenantId"];
            serviceUrl = configuration["serviceUrl"];
            tenantDomainName = configuration["tenantDomainName"];
            stations = configuration["Stations"];
        }
        private async Task<MessagingExtensionActionResponse> CreateCalendarEvent(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            //TeamChannelDetails objChannelDetails = new TeamChannelDetails();
            //if (turnContext.Activity.TeamsGetChannelId() != null)
            //{
            //    objChannelDetails = SimpleTeamsOperations.GetTeamChannelDetails(turnContext, cancellationToken).Result;
            //}

            var magicCode = string.Empty;
            var state = (turnContext.Activity.Value as Newtonsoft.Json.Linq.JObject).Value<string>("state");
            if (!string.IsNullOrEmpty(state))
            {
                int parsed = 0;
                if (int.TryParse(state, out parsed))
                {
                    magicCode = parsed.ToString();
                }
            }

            var tokenResponse = await (turnContext.Adapter as IUserTokenProvider).GetUserTokenAsync(turnContext, _connectionName, magicCode, cancellationToken: cancellationToken);

            if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.Token))
            {
                // There is no token, so the user has not signed in yet.
                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
                var signInLink = await (turnContext.Adapter as IUserTokenProvider).GetOauthSignInLinkAsync(turnContext, _connectionName, cancellationToken);
                return new MessagingExtensionActionResponse
                {
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Type = "auth",
                        SuggestedActions = new MessagingExtensionSuggestedAction
                        {
                            Actions = new List<CardAction>
                                {
                                    new CardAction
                                    {
                                        Type = ActionTypes.OpenUrl,
                                        Value = signInLink,
                                        Title = "Bot Service OAuth",
                                    },
                                },
                        },
                    },
                };
            }
            var accessToken = tokenResponse.Token;
            if (accessToken != null || !string.IsNullOrEmpty(accessToken))
            {
                try
                {

                    var client = new SimpleGraphClient(accessToken);
                    var type = ((JObject)action.Data)["Type"]?.ToString();
                    //var url = turnContext.Activity.Value.ToString();
                    //JObject jsonUrl = JObject.Parse(url);
                    //var link = jsonUrl["messagePayload"]["linkToMessage"];
                    if (type == "CalendarEvent")
                    {

                        //var username = turnContext.Activity.From.AadObjectId;
                        var eventTitle = ((JObject)action.Data)["Title"]?.ToString(); 
                        var eventDetails = ((JObject)action.Data)["Details"]?.ToString(); 
                        var eventStartDt = ((JObject)action.Data)["StartDt"]?.ToString();
                        var eventStartTime = ((JObject)action.Data)["StartTime"]?.ToString();
                        var eventDurationMins = ((JObject)action.Data)["EventDuration"]?.ToString();

                        eventStartDt = eventStartDt + "T" + eventStartTime + ":00";

                        CultureInfo culture = new CultureInfo("en-US");
                        DateTime tempDate = Convert.ToDateTime(eventStartDt, culture);
                        DateTime eventDateStart = DateTime.Now;
                        DateTime eventDateEnd = tempDate.AddMinutes(Int16.Parse(eventDurationMins));
                        var eventEndDate = eventDateEnd.ToString();

                        var eventResponse = await client.CreateOutlookEventAsync(eventTitle, eventStartDt, eventEndDate, eventDetails);

                        var eventUrl = eventResponse.WebLink;

                        eventUrl = "https://teams.microsoft.com/_?culture=en-us&country=US&lm=deeplink&lmsrc=homePageWeb&cmpid=WebSignIn#/scheduling-form/?isBroadcast=false&eventId=" + eventResponse.Id + "&opener=1&providerType=0&navCtx=event-card-peek&calendarType=User";
                      
                        string ackCalendarCardPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Cards", "AckCardCalendarEvents.json");
                        string ackCalendarCardJson = System.IO.File.ReadAllText(ackCalendarCardPath, Encoding.UTF8);

                        ackCalendarCardJson = ackCalendarCardJson.Replace("replaceUrl", eventUrl ?? "", true,
                            culture: CultureInfo.InvariantCulture);
                        ackCalendarCardJson = ackCalendarCardJson.Replace("ReplaceTitel", eventResponse.Subject.ToString() ?? "", true,
                            culture: CultureInfo.InvariantCulture);

                        var card = AdaptiveCard.FromJson(ackCalendarCardJson);
                        Microsoft.Bot.Schema.Attachment attachment = new Microsoft.Bot.Schema.Attachment()
                        {
                            ContentType = AdaptiveCard.ContentType,
                            Content = card.Card
                        };

                        //IMessageActivity cardMsg = MessageFactory.Attachment(attachment);
                        //await turnContext.SendActivityAsync(cardMsg, cancellationToken);

                        return await Task.FromResult(new MessagingExtensionActionResponse
                        {
                            Task = new TaskModuleContinueResponse
                            {
                                Value = new TaskModuleTaskInfo
                                {
                                    Card = attachment,
                                    Height = 180,
                                    Width = 480,
                                    Title = "Event Creation",
                                },
                            },
                        });
                    }
                    return null;
                }

                catch (Exception ex)
                { throw ex; }

            }
            return null;
        }

        // Will be called after OnTeamsMessagingExtensionFetchTaskAsync when user has entered all data in the Messaging Extension Adaptive Card
        private async Task<MessagingExtensionActionResponse> CreateToDoTaskModule(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            //TeamChannelDetails objChannelDetails = new TeamChannelDetails();
            //if (turnContext.Activity.TeamsGetChannelId() != null)
            //{
            //    objChannelDetails = SimpleTeamsOperations.GetTeamChannelDetails(turnContext, cancellationToken).Result;
            //}

            var magicCode = string.Empty;
            var state = (turnContext.Activity.Value as Newtonsoft.Json.Linq.JObject).Value<string>("state");
            if (!string.IsNullOrEmpty(state))
            {
                int parsed = 0;
                if (int.TryParse(state, out parsed))
                {
                    magicCode = parsed.ToString();
                }
            }

            var tokenResponse = await (turnContext.Adapter as IUserTokenProvider).GetUserTokenAsync(turnContext, _connectionName, magicCode, cancellationToken: cancellationToken);

            if (tokenResponse == null || string.IsNullOrEmpty(tokenResponse.Token))
            {
                // There is no token, so the user has not signed in yet.
                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
                var signInLink = await (turnContext.Adapter as IUserTokenProvider).GetOauthSignInLinkAsync(turnContext, _connectionName, cancellationToken);
                return new MessagingExtensionActionResponse
                {
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Type = "auth",
                        SuggestedActions = new MessagingExtensionSuggestedAction
                        {
                            Actions = new List<CardAction>
                                {
                                    new CardAction
                                    {
                                        Type = ActionTypes.OpenUrl,
                                        Value = signInLink,
                                        Title = "Bot Service OAuth",
                                    },
                                },
                        },
                    },
                };
            }
            var accessToken = tokenResponse.Token;
            if (accessToken != null || !string.IsNullOrEmpty(accessToken))
            {
                try
                {
                    
                    var client = new SimpleGraphClient(accessToken);
                    var type = ((JObject)action.Data)["Type"]?.ToString();
                    //var url = turnContext.Activity.Value.ToString();
                    //JObject jsonUrl = JObject.Parse(url);
                    //var link = jsonUrl["messagePayload"]["linkToMessage"];
                    if (type == "ToDoTask")
                    {
                        
                        var username = turnContext.Activity.From.AadObjectId;
                        var taskTitle = ((JObject)action.Data)["Title"]?.ToString();
                        var taskStartDate = ((JObject)action.Data)["StartDate"]?.ToString();
                        var taskDueDate = ((JObject)action.Data)["DueDate"]?.ToString();
                        
                        var itemBody = new Beta.ItemBody();
                        var ToDoResponse = await client.CreateOutlookTaskAsync(taskTitle, taskStartDate, taskDueDate, itemBody);

                        var taskUrl = "https://to-do.office.com/tasks/id/" + ToDoResponse.Id.ToString() + "/details";
                                                
                        string taskCardPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Cards", "todoCardTeams.json");
                        string taskCardJson = System.IO.File.ReadAllText(taskCardPath, Encoding.UTF8);

                        taskCardJson = taskCardJson.Replace("replaceUrl", taskUrl ?? "", true,
                            culture: CultureInfo.InvariantCulture);
                        taskCardJson = taskCardJson.Replace("ReplaceTitel", ToDoResponse.Subject.ToString() ?? "", true,
                            culture: CultureInfo.InvariantCulture);
                    
                        var card = AdaptiveCard.FromJson(taskCardJson);
                        Microsoft.Bot.Schema.Attachment attachment = new Microsoft.Bot.Schema.Attachment()
                        {
                            ContentType = AdaptiveCard.ContentType,
                            Content = card.Card
                        };

                        IMessageActivity cardMsg = MessageFactory.Attachment(attachment);
                        await turnContext.SendActivityAsync(cardMsg, cancellationToken);
                        
                        return await Task.FromResult(new MessagingExtensionActionResponse
                        {
                            Task = new TaskModuleContinueResponse
                            {
                                Value = new TaskModuleTaskInfo
                                {
                                    Card = attachment ,
                                    Height = 180,
                                    Width = 480,
                                    Title = "Task Creation",
                                },
                            },
                        });                      
                    }
                    return null;
                }

                catch (Exception ex)
                { throw ex; }

            }
            return null;
        }

         protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            // check if we previousy requested to install the bot - if true we will present the messaging extension
            if (action.Data.ToString().Contains("justInTimeInstall"))
            {
                return await OnTeamsMessagingExtensionFetchTaskAsync(turnContext, action, cancellationToken);
            }
            else
            {
                switch (action.CommandId)
                {
                    case "CreateCalendarEvent":
                        return await CreateCalendarEvent(turnContext, action, cancellationToken);
                    
                    default:
                        throw new NotImplementedException($"Invalid CommandId: {action.CommandId}");
                }
            }
        }

        // Will be called when user triggers messaging extension
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            AdaptiveCard adaptiveCardEditor;
            try
            {

                 // create task action
                var cardModel = new CardDataModel();
                if (action.CommandId.Equals("CreateCalendarEvent"))
                {
                    var taskTitle = "Enter Task Title"; ;
                    if(action.MessagePayload.Body.Content.ToString().Contains("attachment id"))
                    {
                        var temp = action.MessagePayload.Attachments[0].Content.ToString();
                        var objAttachment = JsonConvert.DeserializeObject<CCCardAttachmentModel>(temp);
                        taskTitle = objAttachment.body[0].text;
                    }
                    else 
                    {
                        taskTitle = action.MessagePayload.Body.Content.ToString();
                    }                    
                    cardModel.TaskTitle = taskTitle;                
                }
                adaptiveCardEditor = AdaptiveCardHelper.CreateCardCalendarEventInputs(cardModel);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return await Task.FromResult(new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = new Microsoft.Bot.Schema.Attachment { Content = adaptiveCardEditor, ContentType = AdaptiveCard.ContentType, },
                        Height = 350,
                        Width = 400,
                        Title = "Event Creation",
                    },
                },
            });
            //Needs to be replaced with OAuth Prompt

        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                var replyText = $"Echo From CollabEventBOT: {turnContext.Activity.Text}";
                await turnContext.SendActivityAsync(MessageFactory.Text(replyText, replyText), cancellationToken);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var welcomeText = "Hello and welcome to Collab Event BOT!";
            foreach (var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(welcomeText, welcomeText), cancellationToken);
                }
            }
        }


    }
}
