// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, CardFactory, TeamsInfo } = require('botbuilder');

class TeamsMessagingExtensionsActionBot extends TeamsActivityHandler {
    async handleTeamsMessagingExtensionFetchTask(context, action) {
        console.log('Context In Fetch Task Block ' + JSON.stringify(context));
        console.log('Action In Fetch Task Block ' + JSON.stringify(action));

        try {
            const teamDetails = await TeamsInfo.getTeamDetails(context);
            console.log('Team details from Teams Info class ' + JSON.stringify(teamDetails));
            const members = await TeamsInfo.getMembers(context);
            console.log('Members details from Teams Info class ' + JSON.stringify(members));

            const peopleArray = [];
            for (let i = 0; i < teamDetails.memberCount; i++) {
                const people = {};
                people.title = members[i].givenName;
                people.value = members[i].email;
                peopleArray.push(people);
            }
            console.log('Array ' + JSON.stringify(peopleArray));
            const IFCard = CardFactory.adaptiveCard({
                type: 'AdaptiveCard',
                body: [
                    {
                        type: 'Container',
                        style: 'emphasis',
                        items: [
                            {
                                type: 'TextBlock',
                                text: 'Instant Feedback',
                                wrap: true
                            }
                        ],
                        padding: 'Default'
                    },
                    {
                        type: 'Container',
                        id: '885220a9-5ab1-95dd-5b66-20f42c452fa9',
                        padding: 'Default',
                        items: [
                            {
                                type: 'TextBlock',
                                weight: 'Bolder',
                                text: 'Whom you want to give feedback?',
                                wrap: true
                            }
                        ],
                        separator: true,
                        spacing: 'None'
                    },
                    {
                        type: 'Container',
                        id: '10017c5a-5ee9-46c5-537a-bdd9ab61225c',
                        padding: {
                            top: 'None',
                            bottom: 'Default',
                            left: 'Default',
                            right: 'Default'
                        },
                        items: [
                            {
                                type: 'Input.ChoiceSet',
                                id: 'options',
                                spacing: 'None',
                                placeholder: 'Placeholder text',
                                choices: peopleArray,
                                style: 'expanded'
                            }
                        ],
                        spacing: 'None'
                    }
                ],
                actions: [
                    {
                        type: 'Action.Submit',
                        id: 'submit',
                        title: 'Submit',
                        data: {
                            action: 'personselector'
                        }
                    }
                ],
                $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                version: '1.0',
                padding: 'None'
            });

            return {
                task: {
                    type: 'continue',
                    value: {
                        card: IFCard,
                        title: 'Instant Feedback',
                        height: 600,
                        width: 700
                    }
                }
            };
        } catch (e) {
            console.log('Error in fetching TeamsInfo ' + e);
        }
    }

    async handleTeamsMessagingExtensionSubmitAction(context, action) {
        console.log('ContextData in Submit action' + JSON.stringify(context));

        const request = require('request-promise');
        try {
            const member = await TeamsInfo.getMember(context, context.activity.from.aadObjectId);
            console.log('Member detail from Teams Info class ' + JSON.stringify(member));

            const options = {
                url: 'https://api.if-staging.haufe.com/api/v1/feedback/initiate-feedback',
                json: true,
                method: 'POST',
                headers: {
                    Authorization: 'Bearer eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiJ4eCIsImlzcyI6ImhhdWZlLmlmIiwic3ViIjoiaGF1ZmUiLCJwZXJzb25JZCI6IjAiLCJyb2xlcyI6IlJPTEVfUFVCTElDX1VTRVIiLCJ0ZW5hbnRJZCI6IjAiLCJpYXQiOjE1OTg0MjY0ODAsImV4cCI6MTYwMzYxMDQ4MH0.MAZB2gxYTFpM-OnzLMgKe0wjxwdHrEaxlRIs4JiQAps',
                    'x-api-key': 'BNOuxbJcr24g6a6Jij0LA8R9CPsq9v4C9aAxeCMX'
                },
                body: {
                    senderEmail: member.email,
                    receiverEmail: context.activity.value.data.options,
                    platform: 'teams'
                }
            };
            const IFResponse = await request(options);
            console.log('Response from IF API ' + JSON.stringify(IFResponse));

            return {
                task: {
                    type: 'continue',
                    value: {
                        url: IFResponse.webUrl,
                        title: 'Instant Feedback',
                        height: 600,
                        width: 700
                    }
                }
            };
        } catch (e) {
            console.log('Error in Submit action ' + e);
        }
    }
}

module.exports.TeamsMessagingExtensionsActionBot = TeamsMessagingExtensionsActionBot;
