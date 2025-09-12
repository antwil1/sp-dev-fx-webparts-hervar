import { SPFI } from "@pnp/sp";
import { LogHelper } from "../helpers/LogHelper";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/taxonomy";
import { ITermInfo } from "@pnp/sp/taxonomy";
import "@pnp/sp/items/get-all";
import "@pnp/sp/search";
import { ISearchQuery, ISearchResult, SearchResults } from "@pnp/sp/search";

export interface ISearchPage {
  items: ISearchResult[];
  total: number;
}

class SPService {
  private static _sp: SPFI;

  public static Init(sp: SPFI): void {
    this._sp = sp;
    LogHelper.info("SPService", "constructor", "PnP SP context initialised");
  }

  public static getListItemsAsync = async (listName: string): Promise<any> => {
    try {
      const items: any = await this._sp.web.lists
        .getByTitle(listName)
        .items.select("*", "ID", "Title")
        .getAll();
      return items;
    } catch (err) {
      LogHelper.error("SPService", "getListItemsAsync", err);
      return null;
    }
  };

  public static getAllTermsByTermSet = async (termSetGuid: string): Promise<any> => {
    try {
      const terms: ITermInfo[] = await this._sp.termStore.sets.getById(termSetGuid).terms();
      return terms;
    } catch (err) {
      LogHelper.error("SPService", "getAllTermsByTermSet", err);
      return null;
    }
  };

  /**
   * @param queryTemplate  QueryTemplate för filter (innehåller managedPropertyName för filter)
   * @param filterManagedPropertyName  MP som används i filtret (TaxID/GUID-variant)
   * @param displayManagedPropertyName MP som ska visas (label/refiner, t.ex. RefinableStringXX)
   * @param page vilken sida som ska hämtas
   * @param pageSize antal objekt per sida
   */
  public static getSearchResults = async (
    queryTemplate: string,
    filterManagedPropertyName: string,
    displayManagedPropertyName?: string,
    page: number = 1,
    pageSize: number = 12
  ): Promise<ISearchPage> => {
    try {
      const selectProps = [
        "Description",
        "DocId",
        "Author",
        "AuthorOWSUSER",
        "Path",
        "NormUniqueID",
        "PictureThumbnailURL",
        "PromotedState",
        "O3CSortableTitle",
        "Title",
        filterManagedPropertyName,
      ];

      if (displayManagedPropertyName) {
        selectProps.push(displayManagedPropertyName);
      }

      const results: SearchResults = await this._sp.search(<ISearchQuery>{
        QueryTemplate: queryTemplate,
        Querytext: "*",
        RowLimit: pageSize,
        StartRow: (page - 1) * pageSize,
        EnableInterleaving: true,
        TrimDuplicates: false,
        SelectProperties: selectProps,
      });

      const items = (results.PrimarySearchResults || []) as unknown as ISearchResult[];
      const total =
        (results as any)?.TotalRows ??
        (results as any)?.RawSearchResults?.PrimaryQueryResult?.RelevantResults?.TotalRows ??
        items.length;

      return { items, total };
    } catch (err) {
      LogHelper.error("SPService", "getSearchResults", err);
      return { items: [], total: 0 };
    }
  };
}
export default SPService;
