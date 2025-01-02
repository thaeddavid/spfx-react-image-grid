import * as React from 'react';
import styles from './ShowcaseGrid.module.scss';
import type { IShowcaseGridProps } from './IShowcaseGridProps';
import { Icon } from '@fluentui/react';

export default class ShowcaseGrid extends React.Component<IShowcaseGridProps, {}> {
  private maxImageWidth = 1000; // Maximum width for resampled images
  private maxImageHeight = 800; // Maximum height for resampled images

  private resampleImage(url: string): Promise<string> {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = "anonymous";  // Required for SharePoint URLs
      
      img.onload = () => {
        // Only resample if image is larger than our max dimensions
        if (img.width <= this.maxImageWidth && img.height <= this.maxImageHeight) {
          resolve(url);
          return;
        }

        // Calculate new dimensions maintaining aspect ratio
        let newWidth = img.width;
        let newHeight = img.height;
        
        if (newWidth > this.maxImageWidth) {
          newHeight = Math.round(newHeight * (this.maxImageWidth / newWidth));
          newWidth = this.maxImageWidth;
        }
        
        if (newHeight > this.maxImageHeight) {
          newWidth = Math.round(newWidth * (this.maxImageHeight / newHeight));
          newHeight = this.maxImageHeight;
        }

        // Create canvas for resampling
        const canvas = document.createElement('canvas');
        canvas.width = newWidth;
        canvas.height = newHeight;
        
        // Draw and resample image
        const ctx = canvas.getContext('2d');
        if (!ctx) return url;
        ctx.drawImage(img, 0, 0, newWidth, newHeight);
        
        // Convert to data URL
        resolve(canvas.toDataURL('image/jpeg', 0.85));
      };

      img.onerror = () => resolve(url); // Fallback to original URL if loading fails

      img.src = url;
    });
  }

  // Cache for resampled images
  private imageCache: { [url: string]: string } = {};

  private async getResampledImage(url: string): Promise<string> {
    if (!url) return '';
    if (this.imageCache[url]) return this.imageCache[url];
    
    try {
      const resampledUrl = await this.resampleImage(url);
      this.imageCache[url] = resampledUrl;
      return resampledUrl;
    } catch (error) {
      console.error('Error resampling image:', error);
      return url;
    }
  }

  public state = {
    resampledImages: {} as { [url: string]: string }
  };

  public componentDidMount() {
    this.loadImages();
  }

  public componentDidUpdate(prevProps: IShowcaseGridProps) {
    if (prevProps.gridItems !== this.props.gridItems) {
      this.loadImages();
    }
  }

  private async loadImages() {
    const { gridItems } = this.props;
    const resampledImages = { ...this.state.resampledImages };

    for (const item of gridItems || []) {
      if (item?.imageUrl?.fileAbsoluteUrl && !resampledImages[item.imageUrl.fileAbsoluteUrl]) {
        resampledImages[item.imageUrl.fileAbsoluteUrl] = await this.getResampledImage(item.imageUrl.fileAbsoluteUrl);
      }
    }

    this.setState({ resampledImages });
  }

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
    const { gridItems, columns, rows } = this.props;
    const maxItems = columns * rows;

    if (!this.hasContent(gridItems)) {
      return (
        <div className={styles.showcaseGrid}>
          <div className={styles.placeholder}>
            <Icon iconName="GridViewMedium" className={styles.placeholderIcon} />
            <h2>No Content Added Yet</h2>
            <p>Configure this web part by clicking the edit button and adding images and content in the properties panel.</p>
          </div>
        </div>
      );
    }

    const style = {
      '--grid-columns': columns
    } as React.CSSProperties;

    return (
      <div className={styles.showcaseGrid}>
        <div className={styles.gridContainer} style={style}>
          {gridItems?.slice(0, maxItems).map((item, index) => (
            <div key={index} className={styles.gridItem}>
              <div className={styles.imageContainer}>
                <img 
                  src={this.state.resampledImages[item.imageUrl.fileAbsoluteUrl] || item.imageUrl.fileAbsoluteUrl}
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