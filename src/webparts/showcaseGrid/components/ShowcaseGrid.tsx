import * as React from 'react';
import styles from './ShowcaseGrid.module.scss';
import type { IShowcaseGridProps } from './IShowcaseGridProps';

export default class ShowcaseGrid extends React.Component<IShowcaseGridProps, {}> {
  private createMarkup(html: string) {
    return { __html: html };
  }

  public render(): React.ReactElement<IShowcaseGridProps> {
    const { gridItems } = this.props;

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
