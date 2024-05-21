import * as React from 'react';
import { ISpFxProps, ISpFxStates } from './ISpFxProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { sp } from 'sp-pnp-js';
import "../../../css/bootstrap.min.css";
import "../../../css/custom.css";
import "../../../css/responsive.css";
import serviceAPI from '../../../APIs/APISCalls';
import styles from '../../../extensions/customHeader/loc/CustomHeader.module.scss';

const SpFx: React.FC<ISpFxProps> = (props: ISpFxProps) => {
  const [state, setState] = React.useState<ISpFxStates>({
    AllItems: [],
    IsEdit: false,
    rootweb: "",
    description: "",
    primarySystemAccount: "",
    recordType: "",
    parantroom: "",
    owner: "",
    restrictions: "",
    RoomDetailsId: 0,
    Team: [],
    disableEdit: false,
    Documents: [],
    showMoreTeam: false,
    showEdit: false,
  });

  React.useEffect(() => {
    const logo = document.getElementById('O365_AppName')
    console.log(logo, 'logo');
    //replace logo inner html span with new content
    if (logo) {
      logo.innerHTML = `
        <span class="${styles.brandtext}">MyGlobalFoundries</span>
        <span class="${styles.sharepointtext}">SharePoint</span>`;
    }

    //select a conpnent which starts with id
    const TitlePaert = document.querySelector('[id^="vpc_WebPart"]');
    //Hide the title part
    if(TitlePaert){
      TitlePaert.innerHTML = '';
    }

    const comment= document.querySelector('CommentsWrapper');
    if(comment){
      comment.innerHTML = '';
    } 
    

    SPComponentLoader.loadCss(
      "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css"
    );
    SPComponentLoader.loadCss(
      "https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css"
    );

    sp.web.get().then((item: any) => {
      console.log("items", item)
      console.log(item.Title)
      setState(prevState => ({
        ...prevState,
        rootweb: item.Title
      }));
    });

    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams?.get('Mode') == 'Edit') {
      setState(prevState => ({
        ...prevState,
        showEdit: true
      }));
    }
    getItems();
  }, []);
  React.useEffect(() => {
    console.log("Context", props.context);
    if (state.Team?.length == 3) {

      setState(prevState => ({
        ...prevState,
        Team: state.TeamCache
      })
      )
    }
    else if (state.IsEdit) {
      setState(prevState => ({
        ...prevState,
        Team: state.TeamCache
      })
      )
    }
    else {
      setState(prevState => ({
        ...prevState,
        Team: state.TeamCache?.slice(0, 3)
      })
      )
    }
  },
    [state.showMoreTeam, state.IsEdit]);

  const getItems = async () => {
    let items = await serviceAPI.getListItems("RoomDetails",
      props?.context?.pageContext?.web?.absoluteUrl,
      [
        "description",
        "primarySystemAccount",
        "recordType",
        "parantroom",
        "owner",
        "restrictions",
        "Id"
      ]);
    console.log("items", items);
    let Team = await serviceAPI.getListItems("Team",
      props?.context?.pageContext?.web?.absoluteUrl,
      ["*",
        "Id",
        "Title",
        "Description",
        "Role",
        "Image",
      ]);
    console.log("Team",
    );
    let Links = await serviceAPI.getListItems("Links",
      props?.context?.pageContext?.web?.absoluteUrl,
      ["*",
        "Id",
        "Name",
        "Link",
        "Group",
      ]);

    let Docuemnts = await serviceAPI.getDocumentsFromLibrary("Shared Documents",
      props?.context?.pageContext?.web?.absoluteUrl,
      ["*",
        "Id",
        "Name",
        "ModifiedBy/Title",
        "Modified",
        "Created",
        "ServerRelativeUrl",
      ],
      [
        "ModifiedBy",
      ],

    );
    let DocCats = await serviceAPI.getDocumentLibraryItems("Documents", props?.context?.pageContext?.web?.absoluteUrl,
      ["*",
        "Id",
        "Title",
        "Category"
      ]
    );

    setState(prevState => ({
      ...prevState,
      description: items[0]?.description,
      primarySystemAccount: items[0]?.primarySystemAccount,
      recordType: items[0]?.recordType,
      parantroom: items[0]?.parantroom,
      owner: items[0]?.owner,
      restrictions: items[0]?.restrictions,
      RoomDetailsId: items[0]?.Id,
      Team: Team?.slice(0, 3)?.map((item: any) => ({
        Id: item.Id,
        Name: item.Name,
        Description: item?.Description,
        Role: item?.Role,
        Image: item?.Image?.Url,
        IsUPdated: false
      })),
      TeamCache: Team?.map((item: any) => ({
        Id: item.Id,
        Name: item.Name,
        Description: item?.Description,
        Role: item?.Role,
        Image: item?.Image?.Url,
        IsUPdated: false
      })),
      Links: Links?.map((item: any) => ({
        Id: item?.Id,
        Name: item?.Name,
        Link: item?.Link,
        Group: item?.Group,
        IsUPdated: false
      })),
      LinksCache: Links?.map((item: any) => ({
        Id: item.Id,
        Name: item.Name,
        Link: item.Link,
        Group: item.Group,
        IsUPdated: false
      })),
      Documents: Docuemnts?.slice(0, 3)?.map((item: any, index) => ({
        Id: item?.Id,
        Name: item?.Name,
        ModifiedBy: item?.ModifiedBy?.Title,
        Modified: item?.TimeLastModified?.split('T')?.[0],
        Link: item?.ServerRelativeUrl,
        Category: DocCats[index]?.Category

      }))
    }));
  }

  const saveData = async () => {

    let RoomDetails = {
      description: state?.description,
      primarySystemAccount: state?.primarySystemAccount,
      recordType: state?.recordType,
      parantroom: state?.parantroom,
      owner: state?.owner,
      restrictions: state?.restrictions,
    }
    // state.RoomDetailsId is zero or null then add new item otherwise update it
    if (state.RoomDetailsId === 0 || state.RoomDetailsId === null || state.RoomDetailsId === undefined) {
      console.log("Add RoomDetails", RoomDetails);
      let RoomDetailsId = await serviceAPI.addListItem("RoomDetails", RoomDetails,
        props?.context?.pageContext?.web?.absoluteUrl);
      console.log("RoomDetailsId", RoomDetailsId);
      setState(prevState => ({
        ...prevState,
        RoomDetailsId: RoomDetailsId.Id
      }));
    }
    else {
      console.log("Update RoomDetails", RoomDetails);
      let RoomDetailsId = await serviceAPI.updateListItemById("RoomDetails", state.RoomDetailsId, RoomDetails,
        props?.context?.pageContext?.web?.absoluteUrl);
      console.log("RoomDetailsId", RoomDetailsId);
    }
    let Team = state.Team.filter((item: any) => item.IsUPdated);
    let Links = state.Links.filter((item: any) => item.IsUPdated);
    // if team is in cache then update it otherwise add it
    let TeamCache = state.TeamCache;
    let LinksCache = state.LinksCache;
    for (let i = 0; i < Team?.length; i++) {
      if (TeamCache.filter((item: any) => item.Id === Team[i].Id)?.length > 0
        && Team[i].IsUPdated === true) {
        let TeamId = await serviceAPI.updateListItemById("Team", Team[i].Id, {
          Name: Team[i].Name,
          Description: Team[i].Description,
          Role: Team[i].Role,
          // Image: Team[i].Image

        },
          props?.context?.pageContext?.web?.absoluteUrl);
        await serviceAPI.uploadImage(props?.context?.pageContext?.web?.absoluteUrl, "Team", Team[i].ImageFile, Team[i].Id);
        console.log("TeamId", TeamId);
      } else if (TeamCache.filter((item: any) => item.Id === Team[i].Id)?.length == 0) {
        let TeamId = await serviceAPI.addListItem("Team", {
          Name: Team[i].Name,
          Description: Team[i].Description,
          Role: Team[i].Role,
          // Image: Team[i].Image
        },
          props?.context?.pageContext?.web?.absoluteUrl);
        if (Team[i].ImageFile && Team[i].ImageFile !== null && Team[i].ImageFile !== undefined && TeamId.Id !== undefined && TeamId.Id !== null) {
          await serviceAPI.uploadImage(props?.context?.pageContext?.web?.absoluteUrl, "Team", Team[i].ImageFile, TeamId.Id);
        }
        console.log("TeamId", TeamId);
      }

    }
    for (let i = 0; i < Links?.length; i++) {
      debugger;
      if (LinksCache.filter((item: any) => item.Id === Links[i].Id &&
        Links[i].IsUPdated === true
      )?.length > 0) {
        debugger;
        console.log("Links[i]", Links[i], "updating item", LinksCache);
        let LinkId = await serviceAPI.updateListItemById("Links", Links[i].Id, {
          Name: Links[i].Name,
          Link: Links[i].Link,
          Group: Links[i].Group
        },
          props?.context?.pageContext?.web?.absoluteUrl);
        console.log("LinkId", LinkId);
      } else if (LinksCache.filter((item: any) => item.Id === Links[i].Id)?.length == 0) {
        console.log("Links[i]", Links[i], "adding new item");
        let LinkId = await serviceAPI.addListItem("Links", {
          Name: Links[i].Name,
          Link: Links[i].Link,
          Group: Links[i].Group
        },
          props?.context?.pageContext?.web?.absoluteUrl);
        console.log("LinkId", LinkId);
      }
    }
    await getItems();
    setState(prevState => ({
      ...prevState,
      IsEdit: false,
      disableEdit: false
    }));
  }

  // get query parameter value

  return (
    <>
      <div
        className="fullScreen"
      >
        <div className="right-section">
          {/* <div className="right_top_section">
            MyGlobalFoundries - Customer Team Room: This is a collaboration space for
            GlobalFoundries and customer teams.
          </div> */}
          <div className="right_section_inner">
            <div className="tab-content">
              <div id="home" className="tab-pane active">
                <div className="share_edit_row_link"
                  hidden={
                    !state.showEdit
                  }
                >
                  <a href="javascript:void(0)" className="edit_link"
                    onClick={() => {
                      setState(prevState => ({
                        ...prevState,
                        IsEdit: !state.IsEdit
                      }));
                      {
                        state.IsEdit ? saveData() : null
                      }
                    }}
                  >
                    {
                      state.IsEdit ? "Save" : "Edit"
                    }
                  </a>
                </div>
                <div className="right_tab_cont">
                  <div className="right_sec_left_sec">
                    <div className="room_detail_wrp">
                      <div className="right_sec_left_heading">
                        <h2>Room Details</h2>
                        <a href="javascript:void(0)">See More</a>
                      </div>
                      <ul className="room_detail">
                        <li>PRIMARY ACCOUNT</li>
                        <li>DESCRIPTION</li>
                        <li>RECORD TYPE</li>
                        {
                          state.IsEdit ?
                            <li>
                              <input
                                type="text"
                                value={state.primarySystemAccount}
                                placeholder='PRIMARY ACCOUNT'
                                onChange={(e) => {
                                  setState(prevState => ({
                                    ...prevState,
                                    primarySystemAccount: e.target.value
                                  }));
                                }}
                              />
                            </li> :
                            <li>{
                              state.primarySystemAccount
                            }</li>
                        }
                        {
                          state.IsEdit ?
                            <li>
                              <input
                                type="text"
                                placeholder='DESCRIPTION'
                                value={state.description}
                                onChange={(e) => {
                                  setState(prevState => ({
                                    ...prevState,
                                    description: e.target.value
                                  }));
                                }}
                              />
                            </li> :
                            <li>{
                              state.description
                            }</li>
                        }
                        {
                          state.IsEdit ?
                            <li>
                              <input
                                type="text"
                                placeholder='RECORD TYPE'
                                value={state.recordType}
                                onChange={(e) => {
                                  setState(prevState => ({
                                    ...prevState,
                                    recordType: e.target.value
                                  }));
                                }}
                              />
                            </li> :
                            <li>{
                              state.recordType
                            }</li>
                        }
                        <li>PARENT ROOM</li>
                        <li>OWNER</li>
                        <li></li>
                        {
                          state.IsEdit ?
                            <li

                            >
                              <input
                                type="text"
                                placeholder='PARENT ROOM'
                                value={state.parantroom}
                                onChange={(e) => {
                                  setState(prevState => ({
                                    ...prevState,
                                    parantroom: e.target.value
                                  }));
                                }}
                              />
                            </li> :
                            <li className='file_td_name'>{
                              state.parantroom
                            }</li>
                        }
                        {
                          state.IsEdit ?
                            <li>
                              <input
                                type="text"
                                placeholder='OWNER'
                                value={state.owner}
                                onChange={(e) => {
                                  setState(prevState => ({
                                    ...prevState,
                                    owner: e.target.value
                                  }));
                                }}
                              />
                            </li> :
                            <li>{
                              state.owner
                            }</li>
                        }
                        <li></li>

                      </ul>
                    </div>
                    <div className="room_detail_wrp">
                      <div className="right_sec_left_heading">
                        <h2>Meet the Account Team</h2>
                        <a href="javascript:void(0)"
                          onClick={() => {
                            setState(prevState => ({
                              ...prevState,
                              showMoreTeam: !state.showMoreTeam
                            }));
                          }}
                        >See More</a>
                      </div>
                      <p className='team_descriptin'>
                        {
                          state.IsEdit ?
                            <input
                              type="text"
                              placeholder='Team Description'
                              value={state.restrictions}
                              onChange={(e) => {
                                setState(prevState => ({
                                  ...prevState,
                                  restrictions: e.target.value
                                }));
                              }
                              }
                            /> :
                            state.restrictions
                        }
                      </p>
                      <div className="team_member_wrp">
                        {
                          state.Team?.map((team: any, index) => (
                            <div className="team_member">
                              <div className="team_member_detail">
                                <div className='photo_upload'>
                                  <div className='photo_upload_image'>
                                    <img
                                      src={
                                        team?.Image ?? 'https://www.w3schools.com/howto/img_avatar.png'
                                      }
                                      // alt='Selected'
                                      style={{
                                        maxWidth: '50px',
                                        maxHeight: '50px',
                                        minWidth: '50px',
                                        minHeight: '50px',
                                        borderRadius: '50%',
                                      }}
                                    />
                                  </div>
                                  <div className='photo_upload_text'
                                    hidden={!state.IsEdit}
                                  >
                                    {/* Modified file input */}
                                    <label
                                      htmlFor={team.Name + index}
                                      className='light_blue_primary_btn ml-0 mt-14 edit_button upload-image'
                                      style={{ width: '60px', height: '40px' }}
                                    >
                                      <i className='fa fa-upload' aria-hidden='true'
                                        hidden={!state.IsEdit}
                                      />
                                      <input
                                        hidden={!state.IsEdit}
                                        key={team.Name + index}
                                        id={team.Name + index}
                                        accept='image/*'

                                        // name={}
                                        type='file'
                                        onChange={
                                          (e) => {
                                            setState(prevState => ({
                                              ...prevState,
                                              Team: state.Team?.map((item: any) => {
                                                if (item.Id === team.Id) {

                                                  item.ImageFile = e.target.files[0];
                                                  item.IsUPdated = true;
                                                  item.Image = URL.createObjectURL(e.target.files[0]);
                                                  console.log('ImageUPdated', URL.createObjectURL(e.target.files[0]));
                                                }
                                                return item;
                                              }
                                              )
                                            })
                                            );
                                          }
                                        }
                                        style={{
                                          display: 'none',
                                          width: '1000px',
                                          height: '1000px',
                                        }}
                                      />
                                    </label>
                                  </div>
                                </div>
                                <div className="team_member_txt_de">
                                  <h3>
                                    <span>
                                      {
                                        state.IsEdit ?
                                          <input
                                            type="text"
                                            value={team.Name}
                                            placeholder='Name'
                                            onChange={(e) => {
                                              setState(prevState => ({
                                                ...prevState,
                                                Team: state.Team?.map((item: any) => {
                                                  if (item.Id === team.Id) {
                                                    item.Name = e.target.value;
                                                    item.IsUPdated = true;
                                                  }
                                                  return item;
                                                })
                                              }));
                                            }}
                                          /> :
                                          team.Name
                                      }</span>
                                    {state.IsEdit ?
                                      <input

                                        type="text"
                                        value={team.Role}
                                        placeholder='Role'
                                        onChange={(e) => {
                                          setState(prevState => ({
                                            ...prevState,
                                            Team: state.Team?.map((item: any) => {
                                              if (item.Id === team.Id) {
                                                item.Role = e.target.value;
                                                item.IsUPdated = true;
                                              }
                                              return item;
                                            })
                                          }));
                                        }}
                                      /> :
                                      team.Role
                                    }
                                  </h3>
                                  <p className='email'>
                                    {
                                      state.IsEdit ?
                                        <input
                                          type="text"
                                          value={team.Description}
                                          placeholder='Email'
                                          onChange={(e) => {
                                            setState(prevState => ({
                                              ...prevState,
                                              Team: state.Team?.map((item: any) => {
                                                if (item.Id === team.Id) {
                                                  item.Description = e.target.value;
                                                  item.IsUPdated = true;
                                                }
                                                return item;
                                              })
                                            }));
                                          }}
                                        /> :
                                        team.Description
                                    }
                                  </p>


                                  {
                                    state.IsEdit ?
                                      <div className="team_member_action"
                                        onClick={() => {
                                          setState(prevState => ({
                                            ...prevState,
                                            Team: state.Team.filter((item: any) => item.Name !== team.Name)
                                          }));
                                        }}
                                      >
                                        <a href="javascript:void(0)">
                                          <i className="fa fa-trash" aria-hidden="true"
                                            onClick={
                                              () => {
                                                setState(prevState => ({
                                                  ...prevState,
                                                  Team: state?.Team?.filter((item: any) => item.Name !== team.Name)
                                                }));
                                                // delete item if it is already present in cache
                                                if (state.TeamCache.filter((item: any) => item.Id === team.Id)?.length > 0) {
                                                  serviceAPI.deleteItemById("Team", team.Id, props?.context?.pageContext?.web?.absoluteUrl);
                                                }
                                              }
                                            }
                                          />
                                        </a>
                                      </div> : null
                                  }

                                </div>
                              </div>
                            </div>
                          ))
                        }

                      </div>

                    </div>
                    {
                      state.IsEdit ?
                        <div className='AddMoreTeam'>
                          <a href="javascript:void(0)">
                            
                            <i className="fa fa-plus" aria-hidden="true"
                              onClick={() => {
                                setState(prevState => ({
                                  ...prevState,
                                  Team: [
                                    ...(state?.Team || []),
                                    {
                                      Name: "",
                                      Description: "",
                                      Role: "",
                                      Image: 'https://www.w3schools.com/howto/img_avatar.png',
                                      IsUPdated: false,
                                      Id: (state?.Team?.length || 0) > 0 ? state?.Team[state.Team?.length - 1].Id + 1 : 1
                                    }
                                  ]
                                }));
                              }}
                            />
                            
                          </a>

                        </div>
                        : null
                    }
                    <div className="room_detail_wrp_res">
                      <div className="room_detail_note">
                        <p>
                          RESTRICTIONS: DO NOT apply any controlled technical information (content or attachment) that is subject to export controls under the Export Administration Regulations (EAR) or International Traffic in Arms Regulations (ITAR). Controlled technical information may be subject to these regulations and licensing if it is peculiarly responsible for achieving or exceeding the control levels of any item listed on the Commerce Control List or the US Munitions List.
                        </p>
                      </div>
                    </div>
                  </div>
                  <div className="mid_right_section">
                    <div className='right_sec_wrap'>
                      <div className="right_sec_left_heading">
                        <h2>Quick Links</h2>
                      </div>
                      <ul className="quick_links_nav">
                        <li>TASKS/ACTIONS</li>
                        {state.Links?.map((link: any) => {
                          if (link.Group === "TASKS/ACTIONS" && !state.IsEdit) {
                            return (
                              <li  >
                                <a href={link.Link}>{link.Name}</a>
                              </li>
                            );
                          }
                          if (link.Group === "TASKS/ACTIONS" && state.IsEdit) {
                            return (
                              <>
                                <li  >
                                  <input
                                    type="text"
                                    value={link.Name}
                                    placeholder='Link Name'
                                    onChange={(e) => {
                                      setState(prevState => ({
                                        ...prevState,
                                        Links: state.Links?.map((item: any) => {
                                          if (item.Name === link.Name) {
                                            item.Name = e.target.value
                                            item.IsUPdated = true;
                                          }
                                          return item;
                                        })
                                      }));
                                    }}
                                  />
                                  <input
                                    type="text"
                                    value={link.Link}
                                    placeholder='Link'
                                    key={'TASKS/ACTIONS' + link.Id}
                                    id={'TASKS/ACTIONS' + link.Id}
                                    onChange={(e) => {
                                      setState(prevState => ({
                                        ...prevState,
                                        Links: state.Links?.map((item: any) => {
                                          if (item.Name === link.Name) {
                                            item.Link = e.target.value;
                                            item.IsUPdated = true;
                                          }
                                          return item;
                                        })
                                      }));
                                    }}
                                  ></input>
                                  <>
                                    <a href="javascript:void(0)">
                                      <i className="fa fa-trash" aria-hidden="true"
                                        onClick={() => {
                                          setState(prevState => ({
                                            ...prevState,
                                            Links: state.Links?.filter((item: any) => item.Name !== link.Name)
                                          }));
                                          // delete item if it is already present in cache
                                          if (state.LinksCache.filter((item: any) => item.Name === link.Name)?.length > 0) {
                                            serviceAPI.deleteItemById("Links", link.Id, props?.context?.pageContext?.web?.absoluteUrl);
                                          }
                                        }
                                        }
                                      />
                                    </a>
                                  </>
                                </li>
                              </>
                            );
                          }
                          return null;
                        })}
                        {
                          state.IsEdit ?
                            <li>
                              <a href="javascript:void(0)"
                                onClick={() => {
                                  setState(prevState => ({
                                    ...prevState,
                                    Links: [...state.Links || [], {
                                      Name: "",
                                      Link: "",
                                      Group: "TASKS/ACTIONS",
                                      Id: state.Links?.length > 0 ? state.Links?.sort((a, b) => a.Id - b.Id)[state.Links?.length - 1].Id + 1 : 1,
                                      IsUPdated: false
                                    }],
                                  }));
                                }}
                                key='TASKS/ACTIONS'
                              >
                                <i className="fa fa-plus" aria-hidden="true" />
                              </a>
                            </li> : null
                        }
                      </ul>
                      <ul className="quick_links_nav">
                        <li>ASSOCIATED ACCOUNTS</li>
                        {state.Links?.map((link: any) => {
                          if (link.Group === "ASSOCIATED ACCOUNTS" && !state.IsEdit) {
                            return (
                              <li  >
                                <a href={link.Link}>{link.Name}</a>
                              </li>
                            );
                          }
                          if (link.Group === "ASSOCIATED ACCOUNTS" && state.IsEdit) {
                            return (
                              <>
                                <li >
                                  <input
                                    type="text"
                                    value={link.Name}
                                    key={'ASSOCIATED ACCOUNTS' + link.Id}
                                    id={'ASSOCIATED ACCOUNTS' + link.Id}
                                    placeholder='Link Name'
                                    onChange={(e) => {
                                      setState(prevState => ({
                                        ...prevState,
                                        Links: state.Links?.map((item: any) => {
                                          if (item.Name === link.Name) {
                                            item.Name = e.target.value;
                                            item.IsUPdated = true;
                                          }
                                          return item;
                                        })
                                      }));
                                    }}
                                  />
                                  <input
                                    type="text"
                                    value={link.Link}
                                    placeholder='Link'
                                    onChange={(e) => {
                                      setState(prevState => ({
                                        ...prevState,
                                        Links: state.Links?.map((item: any) => {
                                          if (item.Name === link.Name) {
                                            item.Link = e.target.value;
                                            item.IsUPdated = true;
                                          }
                                          return item;
                                        })
                                      }));
                                    }}
                                  ></input>
                                  <>
                                    <a href="javascript:void(0)">
                                      <i className="fa fa-trash" aria-hidden="true"
                                        onClick={() => {
                                          setState(prevState => ({
                                            ...prevState,
                                            Links: state.Links?.filter((item: any) => item.Name !== link.Name)
                                          }));
                                          if (state.LinksCache.filter((item: any) => item.Name === link.Name)?.length > 0) {
                                            serviceAPI.deleteItemById("Links", link.Id, props?.context?.pageContext?.web?.absoluteUrl);
                                          }
                                        }}
                                      />
                                    </a>
                                  </>
                                </li>
                              </>
                            );
                          }
                          return null;
                        })}
                        {
                          state.IsEdit ?
                            <li>
                              <a href="javascript:void(0)"
                                key={'ASSOCIATED ACCOUNTS'}
                                id={'ASSOCIATED ACCOUNTS'}
                                onClick={() => {
                                  setState(prevState => ({
                                    ...prevState,
                                    Links:
                                      [...state.Links || [], {
                                        Name: "",
                                        Link: "",
                                        Group: "ASSOCIATED ACCOUNTS",
                                        Id: state.Links?.length > 0 ? state.Links?.sort((a, b) => a.Id - b.Id)[state.Links?.length - 1].Id + 1 : 1,
                                        IsUPdated: false
                                      }]
                                  }));
                                }}
                              >
                                <i className="fa fa-plus" aria-hidden="true" />
                              </a>
                            </li> : null
                        }
                      </ul>
                      <ul className="quick_links_nav">
                        <li>ASSOCIATED TEAM ROOMS</li>
                        {state.Links?.map((link: any, index) => {
                          if (link.Group === "ASSOCIATED TEAM ROOMS" && !state.IsEdit) {
                            return (
                              <li  >
                                <a href={link.Link}>{link.Name}</a>
                              </li>
                            );
                          }
                          if (link.Group === "ASSOCIATED TEAM ROOMS" && state.IsEdit) {
                            return (
                              <>
                                <li>
                                  <input
                                    type="text"
                                    value={link.Name}
                                    placeholder='Link Name'
                                    key={'ASSOCIATED TEAM ROOMS' + link.Id}
                                    id={'ASSOCIATED TEAM ROOMS' + link.Id}
                                    onChange={(e) => {
                                      setState(prevState => ({
                                        ...prevState,
                                        Links: state.Links?.map((item: any) => {
                                          if (item.Name === link.Name) {
                                            item.Name = e.target.value;
                                            item.IsUPdated = true;
                                          }
                                          return item;
                                        })
                                      }));
                                    }}
                                  />
                                  <input
                                    type="text"
                                    value={link.Link}
                                    placeholder='Link'

                                    onChange={(e) => {
                                      setState(prevState => ({
                                        ...prevState,
                                        Links: state.Links?.map((item: any) => {
                                          if (item.Name === link.Name) {
                                            item.Link = e.target.value;
                                            item.IsUPdated = true;
                                          }
                                          return item;
                                        })
                                      }));
                                    }}
                                  ></input>
                                  <>
                                    <a href="javascript:void(0)">
                                      <i className="fa fa-trash" aria-hidden="true"
                                        onClick={() => {
                                          setState(prevState => ({
                                            ...prevState,
                                            Links: state.Links?.filter((item: any) => item.Id !== link.Id)
                                          }));
                                          if (state.LinksCache.filter((item: any) => item.Name === link.Name)?.length > 0) {
                                            serviceAPI.deleteItemById("Links", link.Id, props?.context?.pageContext?.web?.absoluteUrl);
                                          }
                                        }}
                                      />
                                    </a>
                                  </>
                                </li>
                              </>
                            );
                          }
                          return null;
                        })}
                        {
                          state.IsEdit ?
                            <li>
                              <a href="javascript:void(0)"
                                key={'ASSOCIATED TEAM ROOMS'}
                                id={'ASSOCIATED TEAM ROOMS'}
                                onClick={() => {
                                  setState(prevState => ({
                                    ...prevState,
                                    Links:
                                      [...state.Links || [], {
                                        Name: "",
                                        Link: "",
                                        Group: "ASSOCIATED TEAM ROOMS",
                                        Id: state.Links?.length > 0 ? state.Links?.sort((a, b) => a.Id - b.Id)[state.Links?.length - 1].Id + 1 : 1,
                                        IsUPdated: false
                                      }]

                                  }));
                                }}
                              >
                                <i className="fa fa-plus" aria-hidden="true" />
                              </a>
                            </li> : null
                        }
                      </ul>
                    </div>
                  </div>
                </div>
                <div className="res_document_wrp">
                  <div className="right_sec_left_heading">
                    <h2>Recent Documents</h2>
                    <a href="javascript:void(0)"
                      onClick={() => {
                        const url = props?.context?.pageContext?.web?.absoluteUrl + "/Shared%20Documents/Forms/AllItems.aspx";
                        window.open(url);
                      }}
                    >See All</a>
                  </div>
                  <div className="res_document_con">
                    <div className="res_d_table_responsive">
                      <table
                        width="100%"
                        cellPadding={0}
                        cellSpacing={0}
                        // border={0}
                        className="res_document_table"
                      >
                        <tbody>
                          <tr>
                            <th
                            />
                            <th
                            >
                              <span>Name</span>
                            </th>
                            <th
                            >
                              <span>Modified</span>
                            </th>
                            <th
                            >
                              <span>Modified By</span>
                            </th>
                            <th
                            >
                              <span>Category</span>
                            </th>
                          </tr>
                          {
                            state.Documents?.map((item: any) => (
                              <tr>
                                <td>
                                  <i
                                    className="fa fa-file-word-o"
                                    aria-hidden="true"
                                  />
                                </td>
                                <td className="file_td_name"
                                  onClick={() => {
                                    window.open(item.Link);
                                  }}
                                >{item.Name}</td>
                                <td>{item.Modified}</td>
                                <td>{item.ModifiedBy}</td>
                                <td>
                                  <span className={item?.Category?.split(' ')[0]}>{item.Category ?? "--"}</span>
                                </td>
                              </tr>
                            ))
                          }

                          {/* <tr>
                            <td>
                              <i
                                className="fa fa-file-word-o"
                                aria-hidden="true"
                              />
                            </td>
                            <td className="file_td_name">File.dox</td>
                            <td>2 min ago</td>
                            <td>Doe, Patty</td>
                            <td>
                              <span className="weekly_td">Device Weekly</span>
                            </td>
                          </tr>
                          <tr>
                            <td>
                              <i
                                className="fa fa-file-text-o"
                                aria-hidden="true"
                              />
                            </td>
                            <td className="file_td_name">File.text</td>
                            <td>5 hours ago</td>
                            <td>Doe, Patty</td>
                            <td>
                              <span className="special_td">Special Topic</span>
                            </td>
                          </tr>
                          <tr>
                            <td>
                              <i
                                className="fa fa-file-word-o"
                                aria-hidden="true"
                              />
                            </td>
                            <td className="file_td_name">File.dox</td>
                            <td>10 days ago</td>
                            <td>Doe, Patty</td>
                            <td>
                              <span className="cis_w_td">CIS Weekly</span>
                            </td>
                          </tr>  */}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div >
    </>
  )
}

export default SpFx;
