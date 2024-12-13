import * as React from 'react';
import styles from './ShowcaseGrid.module.scss';
import type { IShowcaseGridProps } from './IShowcaseGridProps';

export default class ShowcaseGrid extends React.Component<IShowcaseGridProps, {}> {
  private createMarkup(html: string) {
    return { __html: html };
  }

  private hasContent(gridItems: any[]): boolean {
    return gridItems?.some(item => 
      item?.imageUrl?.fileAbsoluteUrl || 
      item?.title || 
      item?.description || 
      item?.linkUrl
    );
  }

  public render(): React.ReactElement<IShowcaseGridProps> {
    const { gridItems } = this.props;

    if (!this.hasContent(gridItems)) {
      return (
        <div className={styles.showcaseGrid}>
          <div className={styles.placeholder}>
            <h2>No Content Added Yet</h2>
            <p>Configure this web part by clicking the edit button and adding images and content in the properties panel.</p>
          </div>
        </div>
      );
    }

    return (
      <div className={styles.showcaseGrid}>
        <div className={styles.gridContainer}>
          {gridItems?.map((item, index) => (
            <div key={index} className={styles.gridItem}>
              <div className={styles.imageContainer}>
                <img 
                  src={item.imageUrl.fileAbsoluteUrl}
                  alt={item.title || `Grid item ${index + 1}`}
                />
                <div className={styles.overlay}>
                  <h3>{item.title}</h3>
                  <div 
                    className={styles.description}
                    dangerouslySetInnerHTML={this.createMarkup(item.description)}
                  />
                  <a href={item.linkUrl} className={styles.linkButton}>
                    {item.linkText}
                  </a>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }
}