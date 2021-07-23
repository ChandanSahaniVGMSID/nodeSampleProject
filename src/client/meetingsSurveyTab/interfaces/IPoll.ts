export interface IPoll {
    id: string;
    meetingId: string;
    meetingOrganizer: string;
    meetingAttendees: string;
    templateId: string;
    startDateTime: Date;
    endDateTime: Date;
}