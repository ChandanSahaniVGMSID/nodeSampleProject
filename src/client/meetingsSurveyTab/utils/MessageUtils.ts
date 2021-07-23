import { IMessage } from "../interfaces";

export function GetExceptionMessage(e: any): string {
    return !!e && !!e.message ? e.message : e;
}

export function ConcatMessage(messages: Array<IMessage>, message: IMessage): Array<IMessage> {
    if (!message || !message.text) return messages;
    return !!messages ? messages.concat(message) : [message];
}