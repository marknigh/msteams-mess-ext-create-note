const { CardFactory } = require('botbuilder')

function createSectionCard (mySectionNames) {

    const sectionsCard = CardFactory.adaptiveCard({
        version: '1.4',
        type: 'AdaptiveCard',
        body: [
            {
                type: 'Input.ChoiceSet',
                id: 'SectionNames',
                label: 'What Section to insert Page with Note',
                choices: mySectionNames,
                isRequired: true,
                errorMessage: 'Required'
            },
            {
                type: 'Input.Text',
                id: 'Note',
                label: 'What is the note you would like to include',
                isRequired: true,
                errorMessage: 'Required'
            },
        ],
        actions: [
            {
                type: 'Action.Submit',
                title: 'Submit',
                data: {
                    'response': 'createPage'
                }
            }
        ]
    });

    return sectionsCard
}

exports.createSectionCard = createSectionCard