export const Description: string = "Description";

export function GetDescription(enumeration: any, key: any): string {
    if (!enumeration[key]) return "";
    return enumeration[enumeration[key].toString() + Description];
}