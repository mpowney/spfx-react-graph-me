export enum ShowUser {
    CurrentUser = 1,
    SpecifiedUser = 2
}

export default interface IGraphPeopleSettings {
    ShowUser: ShowUser;
    SpecifiedUsername: string;
    FilterOnlyUsers: boolean;
}