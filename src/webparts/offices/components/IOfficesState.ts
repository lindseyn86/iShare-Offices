/* eslint-disable @typescript-eslint/no-explicit-any */
import { IListItem } from "../../../services/SharePoint/IListItem";

export interface IOfficesState {
    items: IListItem[],
    loading: boolean;
    error: string;
}