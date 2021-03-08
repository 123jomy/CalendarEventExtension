// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// Generated with Bot Builder V4 SDK Template for Visual Studio EchoBot v4.9.1
using System.Collections.Generic;
using AdaptiveCards;
using Newtonsoft.Json.Linq;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using System.Threading;
using Newtonsoft.Json;
using System;

using Bot.Builder.Community.Samples.Teams.Models;
using CommonModels;
using Microsoft.Graph;

namespace CustomHelperClasses
{
    public static class AdaptiveCardHelper
    {
        public static AdaptiveCard CreateCardCalendarEventInputs(CardDataModel cardModel)
        {

            // DateTime dueStartDate = new Date
            DateTime dueDateStart = DateTime.Now;
            DateTime dueDateEnd = dueDateStart.AddDays(1);
            var card =  new AdaptiveCard(new AdaptiveSchemaVersion("1.2"))
            {
                Body = new List<AdaptiveElement>
                {new AdaptiveContainer()
                     { Items = new List<AdaptiveElement>()
                        {new AdaptiveTextBlock(){ Id = "Header", Text = "Enter Calendar Event Details",Weight = AdaptiveTextWeight.Bolder },
                        }
                     }
                },
            };

            card.Body.Add(new AdaptiveTextBlock() { Id = "LabelTitle", Text = "Event Title" });
            card.Body.Add(new AdaptiveTextInput() { Id = "Title", Value = cardModel.TaskTitle, IsVisible = true });
            card.Body.Add(new AdaptiveTextBlock() { Id = "LabelDetails", Text = "Event Details" });
            card.Body.Add(new AdaptiveTextInput() { Id = "Details", Placeholder="Enter Event Details", IsVisible = true });
            
            // Row 2
            var colSet = new AdaptiveColumnSet();
            var col1 = new AdaptiveColumn();
            col1.Width = "100";
            col1.Items.Add(new AdaptiveTextBlock() { Id = "LabelStartDt", Text = "Start Date", IsVisible = true });
            var col2 = new AdaptiveColumn();
            col2.Width = "100";
            col2.Items.Add(new AdaptiveDateInput() { Id = "StartDt", Placeholder = "Select Date", Value = dueDateStart.ToString() });            
            var col3 = new AdaptiveColumn();
            col3.Width = "100";
            col3.Items.Add(new AdaptiveTimeInput() { Id = "StartTime", Placeholder = "Select Time", });
            
            colSet.Columns.Add(col1);
            colSet.Columns.Add(col2);
            colSet.Columns.Add(col3);
            card.Body.Add(colSet);

            // Row 3
            var colSet2 = new AdaptiveColumnSet();
            var col21 = new AdaptiveColumn();
            col21.Width = "100";
            col21.Items.Add(new AdaptiveTextBlock() { Id = "LabelDuration", Text = "Event Duration", IsVisible = true });
            var col22 = new AdaptiveColumn();
            col22.Width = "100";
            col22.Items.Add(new AdaptiveChoiceSetInput
                                {   Id = "EventDuration",
                                    Style = AdaptiveChoiceInputStyle.Compact,
                                    Value = "EventDuration",
                                    IsMultiSelect = false,
                                    Choices = new List<AdaptiveChoice>
                                                    {   new AdaptiveChoice() { Title = "15 mins", Value = "15" },
                                                        new AdaptiveChoice() { Title = "30 mins", Value = "30" },
                                                        new AdaptiveChoice() { Title = "45 mins", Value = "45" },
                                                        new AdaptiveChoice() { Title = "1 hr", Value = "60" },
                                                        new AdaptiveChoice() { Title = "2 hr", Value = "120" },
                                                        new AdaptiveChoice() { Title = "3 hr", Value = "180" },
                                                    },
                                }
                            );
            var col23 = new AdaptiveColumn();
            col23.Width = "100";
            col23.Items.Add(new AdaptiveTextBlock() { Id = "LabelDummy2", Text = "", IsVisible = false });
            colSet2.Columns.Add(col21);
            colSet2.Columns.Add(col22);
            colSet2.Columns.Add(col23);
            card.Body.Add(colSet2);

            var actionItem = new AdaptiveActionSet()
            {
                Id = "SubmitCalendarEvent",
                Actions = new List<AdaptiveAction>
                                                {
                                                    new AdaptiveSubmitAction
                                                    {
                                                        Type = AdaptiveSubmitAction.TypeName,
                                                        Title = "Create Calendar Event",
                                                        Id = "SubmitCalendarEvent",
                                                        Data = new JObject { { "Type", "CalendarEvent" } },
                                                    },
                                                },
            };
            card.Body.Add(actionItem);

            return card;
        }
        public static AdaptiveCard CreateCardToDoTaskInputs(CardDataModel cardModel)
        {

            // DateTime dueStartDate = new Date
            DateTime dueDateStart = DateTime.Now;
            DateTime dueDateEnd = dueDateStart.AddDays(1);
            return new AdaptiveCard(new AdaptiveSchemaVersion("1.2"))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveContainer()
                     {
                     Items = new List<AdaptiveElement>()
                        {

                        new AdaptiveTextBlock(){ Id = "Header", Text = "Enter New ToDo Task Details",Weight = AdaptiveTextWeight.Bolder },
                        
                        new AdaptiveTextBlock() { Id = "TitleText", Text = "ToDo Task Title" },
                        new AdaptiveTextInput() { Id = "Title", Value = cardModel.TaskTitle, IsVisible = true },
                        
                        new AdaptiveTextBlock() { Id = "TitleStartDt", Text = "Start Date", IsVisible = true },
                        new AdaptiveDateInput() { Id = "StartDate", Placeholder = "Start date", Value = dueDateStart.ToString(), IsVisible = true},
                        
                        new AdaptiveTextBlock() { Id = "TitleDueDt", Text = "Due Date", IsVisible = true },
                        new AdaptiveDateInput() { Id = "DueDate", Placeholder = "Due date", Value = dueDateEnd.ToString()},

                        new AdaptiveActionSet()
                                            {
                                                Id = "SubmitToDoAction",
                                                Actions = new List<AdaptiveAction>
                                                {
                                                    new AdaptiveSubmitAction
                                                    {
                                                        Type = AdaptiveSubmitAction.TypeName,
                                                        Title = "Create ToDo Task",
                                                        Id = "SubmitToDoTask",                                                        
                                                        Data = new JObject { { "Type", "ToDoTask" } },
                                                    },
                                                },
                                            },
                         
                        }
                     }

                 
                },

            };
        }
               
    }
}