const {
    CardFactory
} = require('botbuilder')

function createNotebookCard (myNoteBookNames) {

    const noteBookCard = CardFactory.adaptiveCard({
        version: '1.4',
        type: 'AdaptiveCard',
        body: [
            {
                type: 'Input.ChoiceSet',
                id: 'NoteBookNames',
                label: 'What NoteBook to Insert Note',
                choices: myNoteBookNames,
                isRequired: true,
                errorMessage: 'Required'
            },
        ],
        actions: [
            {
                type: 'Action.Submit',
                title: 'Get Sections',
                data: {
                    response: 'showSections'
                }
            }
        ]
    });

    return noteBookCard
}

exports.createNotebookCard = createNotebookCard