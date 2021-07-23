export interface IMeeting {
    id: string;
    startDateTime: Date;
    endDateTime: Date;
    subject: string;
    organizer: IMeetingParticipant;
    attendees: Array<IMeetingParticipant>;
}

export interface IMeetingParticipant {
    upn: string;
    id: string;
}