import * as React from "react";
import styles from "./CuratedNews.module.scss";
import { ICuratedNewsProps } from "./ICuratedNewsProps";
import { Card, Col, Row, Space, Spin, Tag } from "antd";
import Meta from "antd/lib/card/Meta";
import SPService from "../../../services/SPService";
import { ISearchResult } from "@pnp/sp/search";
import GraphService from "../../../services/GraphService";
import CachingService from "../../../services/CachingService";
import { ITerm } from "../../preferences/types/Component.Types";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

export const CuratedNews: React.FC<ICuratedNewsProps> = (props) => {
  const {
    extensionName,
    loginName,
    title,
    managedPropertyName, // används för FILTER i sökningen (ofta TaxID/GUID-MP)
    context,
    newsPageLink,
    enableCaching,
  } = props;

  const DISPLAY_PROP = "RefinableString01";
  const [data, setData] = React.useState<ISearchResult[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const preferenceCacheKey = `CuratedNews-UserPreferences-${loginName}`;

  const onConfigure = () => context.propertyPane.open();

  const getUserPreferences = React.useCallback(async () => {
    const cachedData = CachingService.get(preferenceCacheKey);
    if (cachedData !== null) return cachedData;

    const result = await GraphService.GetPreferences(extensionName);
    if (result && result.Tags && result.Tags.length > 0 && enableCaching) {
      CachingService.set(preferenceCacheKey, result.Tags);
    }
    return result?.Tags || [];
  }, [preferenceCacheKey, extensionName, enableCaching]);

  const fetchData = React.useCallback(async () => {
    setLoading(true);
    const tags = await getUserPreferences();

    if (!Array.isArray(tags) || tags.length === 0) {
      setData([]);
      setLoading(false);
      return [];
    }

    const queryTemplate = composeQueryTemplate(tags);
    if (!queryTemplate) {
      setData([]);
      setLoading(false);
      return [];
    }

    const result = await SPService.getSearchResults(queryTemplate, managedPropertyName, DISPLAY_PROP);
    return result;
  }, [getUserPreferences, managedPropertyName, DISPLAY_PROP]);

  React.useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const results = await fetchData();
        if (alive) setData(results ?? []);
      } finally {
        if (alive) setLoading(false);
      }
    })();
    return () => { alive = false; };
  }, [fetchData]);

  if (!extensionName || !managedPropertyName || !newsPageLink) {
    return (
      <Placeholder
        iconName="Edit"
        iconText="Configure your web part"
        description="Please provide the Microsoft Graph open extension name and managed property name."
        buttonLabel="Configure"
        onConfigure={onConfigure}
      />
    );
  }

  return (
    <section>
      <div className={styles["news-container"]}>
        <Spin spinning={loading} tip="Loading...">
          <Card
            title={title}
            headStyle={{ fontSize: "2rem" }}
            extra={<a href={newsPageLink}>Visa alla</a>}
          >
            <Row gutter={16}>
              {data.length > 0 &&
                data.map((newsItem: any) => {
                  const raw: string | undefined =
                    newsItem[DISPLAY_PROP] ?? newsItem[managedPropertyName];

                  const tags: string[] = raw
                    ? raw
                        .split(";")
                        .map(s => (s.includes("|") ? s.split("|")[0] : s))
                        .map(s => s.trim())
                        .filter(Boolean)
                    : [];

                  return (
                    <Col key={newsItem.DocId} span={6}>
                      <Card
                        className={styles.newsCard}
                        hoverable
                        bordered={false}
                        style={{ cursor: "pointer" }}
                        onClick={() => (window.location.href = newsItem.Path)}
                        cover={
                          <img
                            alt={newsItem.Title}
                            src={newsItem.PictureThumbnailURL}
                          />
                        }
                        actions={[
                          <div
                            key={`tags-${newsItem.DocId}`}
                            style={{ width: "100%" }}
                            onClick={(e) => e.stopPropagation()}
                          >
                            <Space size={[8, 8]} wrap className={styles.tags}>
                              {tags.map((tag) => (
                                <Tag key={tag} color="#108ee9">{tag}</Tag>
                              ))}
                            </Space>
                          </div>,
                        ]}
                      >
                        <Meta
                          title={<a href={newsItem.Path}>{newsItem.Title}</a>}
                          description={
                            <>
                              <span className={styles.description}>
                                {newsItem.Description}
                              </span>
                              <div style={{ marginTop: 10 }} />
                            </>
                          }
                        />
                      </Card>
                    </Col>
                  );
                })}
            </Row>
          </Card>
        </Spin>
      </div>
    </section>
  );

  function composeQueryTemplate(tags: ITerm[]) {
    if (!Array.isArray(tags) || tags.length === 0) {
      return null;
    }
    let filterQuery = "";
    if (Array.isArray(tags) && tags.length > 0) {
      const taxValues = `(${tags.map((t) => t.id).join(" OR ")})`;
      filterQuery = `({|${managedPropertyName}:${taxValues}})`;
    }
    const queryTemplate = `{searchTerms} (ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118*) PromotedState=2 ${filterQuery || ""}`;
    return queryTemplate;
  }
};
