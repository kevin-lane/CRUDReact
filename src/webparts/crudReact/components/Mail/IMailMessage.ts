export interface IMailMessage {
    message: {
        subject: string;
        body: {
            contentType: string;
            content: string;
        },
        toRecipients: IRecipient[],
    };
    saveToSentItems: boolean;
}

export interface IRecipient {
    emailAddress: {
        address: string;
    };
}