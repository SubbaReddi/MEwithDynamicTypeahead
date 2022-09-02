// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using TypeaheadSearch.Models;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class TeamsMessagingExtensionsActionPreviewBot : TeamsActivityHandler
    {
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            if (turnContext.Activity.Value != null)
            {
                // This was a message from the card.
                var obj = (JObject)turnContext.Activity.Value;
                var answer = obj["Answer"]?.ToString();
                var choices = obj["Choices"]?.ToString();
                await turnContext.SendActivityAsync(MessageFactory.Text($"{turnContext.Activity.From.Name} answered '{answer}' and chose '{choices}'."), cancellationToken);
            }
            else
            {
                // This is a regular text message.
                await turnContext.SendActivityAsync(MessageFactory.Text($"Hello from the TeamsMessagingExtensionsActionPreviewBot."), cancellationToken);
            }
        }

        protected override Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            //   var adaptiveCardEditor = AdaptiveCardHelper.CreateAdaptiveCardEditor();
            var adaptiveCardEditor = AdaptiveCardHelper.CreateAdaptiveCardAttachment();

            return Task.FromResult(new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        //Card = new Attachment
                        //{
                        //    Content = adaptiveCardEditor,
                        //    ContentType = AdaptiveCard.ContentType,
                        //},
                        Card = adaptiveCardEditor,
                        Height = 450,
                        Width = 500,
                        Title = "Task Module Fetch Example",
                    },
                },
            });
        }

        protected override Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            var exampleData = JsonConvert.DeserializeObject<ExampleData>(action.Data.ToString());

            var adaptiveCard = AdaptiveCardHelper.CreateAdaptiveCard(exampleData);

            // a number of reasonable options here...

            // (1) send a message on a new conversation and return null (only works in group chats and teams)

            // THIS WILL WORK IF THE BOT IS INSTALLED. (GetMembers() will NOT throw if the bot is installed.)

            //var message = MessageFactory.Attachment(new Attachment { ContentType = AdaptiveCard.ContentType, Content = adaptiveCard });
            //var channelId = turnContext.Activity.TeamsGetChannelId();
            //await turnContext.TeamsCreateConversationAsync(channelId, message, cancellationToken);
            //return null;

            // (2) drop the content into the compose window ready for the user to send

            //return new MessagingExtensionActionResponse
            //{
            //    ComposeExtension = new MessagingExtensionResult
            //    {
            //        Type = "result",
            //        AttachmentLayout = "list",
            //        Attachments = new List<MessagingExtensionAttachment>
            //        {
            //            new MessagingExtensionAttachment
            //            {
            //                Content = adaptiveCard,
            //                ContentType = AdaptiveCard.ContentType,
            //            },
            //        },
            //    },
            //};

            // (3) start a preview flow

            return Task.FromResult(new MessagingExtensionActionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "botMessagePreview",
                    ActivityPreview = MessageFactory.Attachment(new Attachment
                    {
                        Content = adaptiveCard,
                        ContentType = AdaptiveCard.ContentType,
                    }) as Activity,
                },
            });
        }

        protected override Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionBotMessagePreviewEditAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            // The data has been returned to the bot in the action structure.
            var activityPreview = action.BotActivityPreview[0];
            var attachmentContent = activityPreview.Attachments[0].Content;
            var previewedCard = JsonConvert.DeserializeObject<AdaptiveCard>(attachmentContent.ToString(), new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            var exampleData = AdaptiveCardHelper.CreateExampleData(previewedCard);

            // This is a preview edit call and so this time we want to re-create the adaptive card editor.
            var adaptiveCardEditor = AdaptiveCardHelper.CreateAdaptiveCardAttachment();

            return Task.FromResult(new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        //Card = new Attachment
                        //{
                        //    Content = adaptiveCardEditor,
                        //    ContentType = AdaptiveCard.ContentType,
                        //},
                        Card = adaptiveCardEditor,
                        Height = 450,
                        Width = 500,
                        Title = "Task Module Fetch Example",
                    },
                },
            });
        }

        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionBotMessagePreviewSendAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            // The data has been returned to the bot in the action structure.
            var activityPreview = action.BotActivityPreview[0];
            var attachmentContent = activityPreview.Attachments[0].Content;
            var previewedCard = JsonConvert.DeserializeObject<AdaptiveCard>(attachmentContent.ToString(), new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            var exampleData = AdaptiveCardHelper.CreateExampleData(previewedCard);

            // This is a send so we are done and we will create the adaptive card editor.
            var adaptiveCard = AdaptiveCardHelper.CreateAdaptiveCard(exampleData);

            var message = MessageFactory.Attachment(new Attachment { ContentType = AdaptiveCard.ContentType, Content = adaptiveCard });

            //User Attribution for Bot messages
            if (exampleData.UserAttributionSelect == "true")
            {
                message.ChannelData = new
                {
                    OnBehalfOf = new[]
                   {
                    new
                       {
                         ItemId = 0,
                         MentionType = "person",
                         Mri = turnContext.Activity.From.Id,
                         DisplayName = turnContext.Activity.From.Name
                    }
                }
                };
            }

            // THIS WILL WORK IF THE BOT IS INSTALLED. (SendActivityAsync will throw if the bot is not installed.)
            await turnContext.SendActivityAsync(message, cancellationToken);

            return null;
        }

        protected override async Task OnTeamsMessagingExtensionCardButtonClickedAsync(ITurnContext<IInvokeActivity> turnContext, JObject obj, CancellationToken cancellationToken)
        {
            // If the adaptive card was added to the compose window (by either the OnTeamsMessagingExtensionSubmitActionAsync or
            // OnTeamsMessagingExtensionBotMessagePreviewSendAsync handler's return values) the submit values will come in here.
            var reply = MessageFactory.Text("OnTeamsMessagingExtensionCardButtonClickedAsync Value: " + JsonConvert.SerializeObject(turnContext.Activity.Value));
            await turnContext.SendActivityAsync(reply, cancellationToken);
        }

        /// <summary>
        ///  Invoked when an invoke activity is received from the connector.
        /// </summary>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<InvokeResponse> OnInvokeActivityAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            InvokeResponse adaptiveCardResponse;
            if (turnContext.Activity.Name == "application/search")
            {
                var searchData = JsonConvert.DeserializeObject<DynamicSearchCard>(turnContext.Activity.Value.ToString());
                var packageResult = JObject.Parse(await (new HttpClient()).GetStringAsync($"https://azuresearch-usnc.nuget.org/query?q=id:{searchData.queryText}&prerelease=true"));
                if (packageResult == null)
                {
                    var searchResponseData = new
                    {
                        type = "application/vnd.microsoft.search.searchResponse"
                    };

                    var jsonString = JsonConvert.SerializeObject(searchResponseData);
                    JObject jsonData = JObject.Parse(jsonString);

                    adaptiveCardResponse = new InvokeResponse()
                    {
                        Status = 204,
                        Body = jsonData
                    };
                }
                else
                {
                    var packages = packageResult["data"].Select(item => (item["id"].ToString(), item["description"].ToString()));
                    var packageList = packages.Select(item => { var obj = new { title = item.Item1, value = item.Item1 + " - " + item.Item2 }; return obj; }).ToList();
                    var searchResponseData = new
                    {
                        type = "application/vnd.microsoft.search.searchResponse",
                        value = new
                        {
                            results = packageList
                        }
                    };

                    var jsonString = JsonConvert.SerializeObject(searchResponseData);
                    JObject jsonData = JObject.Parse(jsonString);

                    adaptiveCardResponse = new InvokeResponse()
                    {
                        Status = 200,
                        Body = jsonData
                    };
                }

                return adaptiveCardResponse;
            }
            else
            {
                return CreateInvokeResponse(await OnTeamsMessagingExtensionFetchTaskAsync(turnContext, SafeCast<MessagingExtensionAction>(turnContext.Activity.Value), cancellationToken).ConfigureAwait(false));
            }

            return null;
        }

        /// <summary>
        /// Safely casts an object to an object of type <typeparamref name="T"/> .
        /// </summary>
        /// <param name="value">The object to be casted.</param>
        /// <returns>The object casted in the new type.</returns>
        private static T SafeCast<T>(object value)
        {
            var obj = value as JObject;
            if (obj == null)
            {
                throw new Exception($"expected type '{value.GetType().Name}'");
            }

            return obj.ToObject<T>();
        }

    }
}
