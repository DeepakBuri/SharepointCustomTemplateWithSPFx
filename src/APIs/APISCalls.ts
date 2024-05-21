// import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IBatchUpdateReq {
  ListName: string;
  IBatchItems: IBatchItems[];
}

export interface IBatchItems {
  Data: any;
  IsAdd: boolean;
  Id: any;
}

//import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { Web, sp } from 'sp-pnp-js';
export default class serviceAPI {

  public static getListItems(
    listName: string,
    siteUrl: string,
    selectedColumns?: string[],
    expandColumns?: string[],
    filterCondition?: string,
    topMax?: number,
    sortingColumn?: string,
    sortingOrder?: boolean,

  ): Promise<any[]> {
    return new Promise<any[]>((resolve, reject): void => {
      let web = new Web(siteUrl);
      let listItems: any;
      if (listName.indexOf("User Information List") > -1)
        listItems = web.lists.getByTitle(listName).items;
      else listItems = web.lists.getByTitle(listName).items;
      if (selectedColumns != undefined && selectedColumns.length > 0)
        listItems = listItems.select(selectedColumns.toString());
      if (expandColumns != undefined && expandColumns.length > 0)
        listItems = listItems.expand(expandColumns.toString());
      if (filterCondition != undefined && filterCondition.length > 0)
        listItems = listItems.filter(filterCondition);
      if (topMax != undefined && topMax > 0) listItems = listItems.top(topMax);
      if (sortingColumn != undefined && sortingColumn.length > 0)
        listItems = listItems.orderBy(
          sortingColumn,
          sortingOrder === false ? false : true
        );
      listItems.get().then(
        (alllistItems: any[]): void => {
          console.log(
            "Success result while getting items from list '" +
            listName +
            "' in method getListItems():"
          );
          console.log(alllistItems);
          resolve(alllistItems);
        },
        (error: any): void => {
          console.log(
            "Exception while getting items from list '" +
            listName +
            "' in method getListItems():"
          );
          reject(error);
        }
      );
    });
  }

  public static addListItem(listName: string, itemObj: any, siteUrl: string): Promise<any> {
    let web = new Web(siteUrl);
    return web.lists
      .getByTitle(listName)
      .items.add(itemObj)
      .then((response) => {
        console.log(
          "Success result while adding item to list '" +
          listName +
          "' in method addListItem():"
        );
        console.log(response.data);
        return response.data;
      })
      .catch((error) => {
        console.log(
          "Exception while adding item to List '" +
          listName +
          "' in method addListItem():"
        );

        throw new Error(error);
      });
  }

  public static batchAddListItem(
    listName: string,
    itemObj: any[],
    siteUrl: string
  ): Promise<any> {
    let web = new Web(siteUrl);
    let list = web.lists.getByTitle(listName);
    return list.getListItemEntityTypeFullName().then((entityTypeFullName) => {
      let batch = web.createBatch();
      itemObj.forEach((value, index, array) => {
        list.items.inBatch(batch).add(value, entityTypeFullName);
      });
      return batch
        .execute()
        .then((response) => {
          console.log(
            "Success result while adding items to list '" +
            listName +
            "' in method batchAddListItem():"
          );
          //console.log(response);
          return response;
        })
        .catch((error) => {
          console.log(
            "Exception while adding items to List '" +
            listName +
            "' in method batchAddListItem():"
          );

          throw new Error(error);
        });
    });
  }

  public static updateListItemById(
    listName: string,
    itemID: number,
    itemObj: any,
    siteUrl: string
  ): Promise<any> {
    let web = new Web(siteUrl);
    return web.lists
      .getByTitle(listName)
      .items.getById(itemID)
      .update(itemObj)
      .then((response) => {
        console.log(
          "Success result while updating item in list '" +
          listName +
          "' in method updateListItemById():"
        );

        return response.data;
      })
      .catch((error) => {
        throw new Error(error);
      });
  }

  public static batchUpdateListItemById(
    listName: string,
    itemIDs: number[],
    itemObj: any[],
    siteUrl: string
  ): Promise<any> {
    let web = new Web(siteUrl);
    let list = web.lists.getByTitle(listName);
    return list.getListItemEntityTypeFullName().then((entityTypeFullName) => {
      let batch = web.createBatch();
      itemIDs.forEach((value, index, array) => {
        list.items
          .getById(value)
          .inBatch(batch)
          .update(itemObj[index], "*", entityTypeFullName);
      });
      return batch
        .execute()
        .then((response) => {
          console.log(
            "Success result while updating items in list '" +
            listName +
            "' in method batchUpdateListItemById():"
          );
          //console.log(response);
          return response;
        })
        .catch((error) => {
          console.log(
            "Exception while updating items in List '" +
            listName +
            "' in method batchUpdateListItemById():"
          );

          throw new Error(error);
        });
    });
  }

  public static async batchInsertUpdateMultipleListItem(
    data: IBatchUpdateReq[],
    siteUrl: string
  ): Promise<any> {
    let web = new Web(siteUrl);
    let batch = web.createBatch();
    for (var i = 0; i < data.length; i++) {
      let list = web.lists.getByTitle(data[i].ListName);
      const entityTypeFullName = await list.getListItemEntityTypeFullName();
      for (var j = 0; j < data[i].IBatchItems.length; j++) {
        if (data[i].IBatchItems[j].IsAdd) {
          list.items
            .inBatch(batch)
            .add(data[i].IBatchItems[j].Data, entityTypeFullName);
        } else {
          list.items
            .getById(data[i].IBatchItems[j].Id)
            .inBatch(batch)
            .update(data[i].IBatchItems[j].Data, "*", entityTypeFullName);
        }
      }
    }
    return batch
      .execute()
      .then((response) => {
        console.log(
          "Success result while adding items to list '" +
          JSON.stringify(data) +
          "' in method batchInsertUpdateMultipleListItem():"
        );
        //console.log(response);
        return response;
      })
      .catch((error) => {
        console.log(
          "Exception while adding items to List '" +
          JSON.stringify(data) +
          "' in method batchInsertUpdateMultipleListItem():"
        );

        throw new Error(error);
      });
  }

  public static deleteItemById(listName: string, id: any, siteUrl: string): Promise<any> {
    let web = new Web(siteUrl);
    return web.lists
      .getByTitle(listName)
      .items.getById(id)
      .delete()
      .then((response) => {
        console.log(
          `Success result while deleting item from list '${listName}' in method deleteItemById():`
        );
        return response;
      })
      .catch((error) => {
        console.log(
          `Exception while deleting item from list '${listName}' in method deleteItemById():`
        );

        throw new Error(error);
      });
  }

  public static getUserProfile(UserName: string, key?: string[]) {
    return sp.profiles
      .getPropertiesFor(`i:0#.f|membership|${UserName}`)
      .then((data) => {
        console.log(
          "Success result while getting User Profile in method getUserProfile()"
        );
        if (!key) return data.UserProfileProperties;
        else {
          let result: any[] = [];
          key.forEach((element) => {
            if (data.UserProfileProperties.some((i: { Key: string; }) => i.Key == element))
              result.push(
                data.UserProfileProperties.filter((i: { Key: string; }) => i.Key == element)[0]
                  .Value
              );
            else result.push(null);
          });
          return result;
        }
      })
      .catch((error) => {
        console.log(
          "Exception while getting User Profile in method getUserProfile()"
        );

      });
  }

  public static addAttachment(
    SiteUrl: string,
    listName: string,
    Id: number,
    fileName: string,
    content: any
  ): Promise<any> {
    return new Promise<any[]>((resolve, reject): void => {
      let web = new Web(
        SiteUrl
      );
      web.lists
        .getByTitle(listName)
        .items.getById(Id)
        .attachmentFiles.add(fileName, content)
        .then(
          (response: any): void => {
            console.log(
              "Success result while uploading attachment to list '" +
              listName +
              "' in method addAttachment():"
            );
            console.log(response);
            resolve(response);
          },
          (error: any): void => {
            console.log(
              "Exception while uploading attachment to list '" +
              listName +
              "' in method addAttachment():"
            );

            reject(error);
          }
        );
    });
  }
  public static getDocumentLibraryItems(
    listName: string,
    siteUrl: string,
    selectedColumns?: string[],
    expandColumns?: string[],
    filterCondition?: string,
    topMax?: number,
    sortingColumn?: string,
    sortingOrder?: boolean,

  ): Promise<any[]> {
    return new Promise<any[]>((resolve, reject): void => {
      let web = new Web(siteUrl);
      let listItems: any;
      listItems = web.lists.getByTitle(listName).items;
      if (selectedColumns != undefined && selectedColumns.length > 0)
        listItems = listItems.select(selectedColumns.toString());
      if (expandColumns != undefined && expandColumns.length > 0)
        listItems = listItems.expand(expandColumns.toString());
      if (filterCondition != undefined && filterCondition.length > 0)
        listItems = listItems.filter(filterCondition);
      if (topMax != undefined && topMax > 0) listItems = listItems.top(topMax);
      if (sortingColumn != undefined && sortingColumn.length > 0)
        listItems = listItems.orderBy(
          sortingColumn,
          sortingOrder === false ? false : true
        );
      listItems.get().then(
        (alllistItems: any[]): void => {
          console.log(
            "Success result while getting items from list '" +
            listName +
            "' in method getListItems():"
          );
          console.log(alllistItems);
          //sort by modified date
          alllistItems.sort((a, b) => new Date(b?.Modified).getTime() - new Date(a?.Modified).getTime());
          resolve(alllistItems);
        },
        (error: any): void => {
          console.log(
            "Exception while getting items from list '" +
            listName +
            "' in method getListItems():"
          );
          reject(error);
        }
      );
    });
  }

  public static uploadImage(
    SiteUrl: string,
    listName: string,
    file: File,
    Id?: number,
  ): Promise<any> {
    return new Promise<any[] | void>(async (resolve, reject): Promise<void> => {
      let web = new Web(SiteUrl);
      if (file == null || file == undefined) {
        reject("Error while uploading file");
        return;
      }
      // upload to the root folder of site assets in this demo
      const assets = await web.lists.ensureSiteAssetsLibrary();
      const fileItem = await assets.rootFolder.files.add(file.name, file, true);
      // file item is not null or undefined then proceed

      // bare minimum; probably you'll want other properties as well
      const img = {
        "Description": "myImage",
        "Url": fileItem.data.ServerRelativeUrl
      };
      console.log(img, 'Image');

      try {
        // update the item, stringify json for image column
        if (Id != 0) {
          await web.lists.getByTitle(listName).items.getById(Id).update({
            Image: img
          });
        }
        // create the item, stringify json for image column
        else {
          await web.lists.getByTitle(listName).items.add({
            Image: img
          });
        }
        resolve(); // resolve the promise if everything is successful
      } catch (error) {
        reject(error); // reject the promise if there is an error
      }

    });
  }

  public static getDocumentsFromLibrary(
    LiberaryName: string,
    SiteUrl: string,
    selectedColumns?: string[],
    expandColumns?: string[],
  ): Promise<any[]> {
    return new Promise<any[]>((resolve, reject): void => {

      let listItems: any;
      let web = new Web(SiteUrl);
      // get All columns from the list the document library

      listItems = web.getFolderByServerRelativeUrl(LiberaryName).files;
      console.log(listItems, 'Files');
      if (selectedColumns != undefined && selectedColumns.length > 0)
        listItems = listItems.select(selectedColumns.toString());
      if (expandColumns != undefined && expandColumns.length > 0)
        listItems = listItems.expand(expandColumns.toString());
      listItems.get().then(
        (alllistItems: any[]): void => {
          console.log(
            "Success result while getting items from list '" +
            LiberaryName +
            "' in method getListItems():"
          );
          //sort by modified date
          console.log(alllistItems, 'All Items');
          alllistItems.sort((a, b) => new Date(b?.TimeLastModified).getTime() - new Date(a?.TimeLastModified).getTime());
          resolve(alllistItems);
        },
        (error: any): void => {
          console.log(
            "Exception while getting items from list '" +
            LiberaryName +
            "' in method getListItems():"
          );
          reject(error);
        }
      );
    });
  }

  // write a method to add a page to site page document liberary
  public static addPageToSitePages(
    SiteUrl: string,
    pageName: string,
  ): Promise<any> {
    return new Promise<any[]>(async (resolve, reject): Promise<void> => {

      let web = new Web(SiteUrl);
      try {
        // if page with name HomePage exists then do nothing else create a page with name HomePage and set it as welcome page
        const items = await web.lists.getByTitle("Site Pages").items.filter(
          "Title eq 'Root'",
        ).get();
        if (items.length === 0) {
          const page = await web.addClientSidePage(pageName, "Root", "Site Pages");
          console.log(
            "Success result while adding page to site pages in method addPageToSitePages():"
          );

          await page.disableComments();
          await page.save();
          await web.rootFolder.update(
            { WelcomePage: "SitePages/" + pageName }
          );
          console.log("Home page set successfully");
        } else {
          console.log('Page already exists');
        }
        // make root folder as welcome page
        await web.rootFolder.update(
          { WelcomePage: "SitePages/" + pageName }
        );
        //add web part to the page by web part id
        // resolve();
      } catch (error) {
        console.log(
          "Exception while adding page to site pages in method addPageToSitePages():",
          error
        );
        reject(error);
      }
    });
  }

}