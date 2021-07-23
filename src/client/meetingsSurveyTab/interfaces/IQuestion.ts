import { QuestionType } from "../constants/QuestionType";

export interface IQuestion {
    id: string;
    title : string;
    type: QuestionType;
    isRequired: boolean;
}