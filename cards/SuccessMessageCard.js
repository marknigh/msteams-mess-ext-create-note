
const { CardFactory } = require('botbuilder')

function createSuccessCard () {

    const successCard = CardFactory.adaptiveCard({
        version: '1.4',
        type: 'AdaptiveCard',
        body: [
            {
                type: 'TextBlock',
                id: 'final',
                Text: 'Done!',
                },
        ],
        actions: [
            {
                type: 'Action.Submit',
                title: 'Done!',
                data: {
                    'response': 'default'
                }
            }
        ]
    });

    return successCard
    
}

exports.createSuccessCard = createSuccessCard
