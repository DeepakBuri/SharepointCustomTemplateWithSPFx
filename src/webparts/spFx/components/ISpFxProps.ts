export interface ISpFxProps {
  description: string;
  context: any | null;
  primarySystemAccount: string;
  recordType: string;
  parantroom: string;
  owner: string;
  restrictions: string;
}

export interface ISpFxStates {
  AllItems: any[];
  rootweb: string;
  IsEdit: boolean;
  description: string;
  primarySystemAccount: string;
  recordType: string;
  parantroom: string;
  owner: string;
  restrictions: string;
  RoomDetailsId: number;
  disableEdit: boolean;
  Team?: {
    Id: number;
    Name: string;
    Description: string;
    Role: string;
    Image: string;
    ImageFile?: File| null;
    IsUPdated: boolean;
  }[],
  TeamCache?:{
    Id: number;
    Name: string;
    Description: string;
    Role: string;
    Image: string;
    IsUPdated: boolean;
  }[],
  Links?: {
    Id: number;
    Name: string;
    Link: string;
    Group: string;
    IsUPdated: boolean;
  }[],
  LinksCache?: {
    Id: number;
    Name: string;
    Link: string;
    Group: string;
    IsUPdated: boolean;
  }[], 
  Documents?: {
    Id: number;
    Name: string;
    Modified: string;
    ModifiedBy: string;
    Link:string
  }[],
  showMoreTeam: boolean;
  showEdit: boolean;
}
