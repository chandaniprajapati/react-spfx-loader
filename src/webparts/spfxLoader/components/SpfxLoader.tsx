import * as React from 'react';
import styles from './SpfxLoader.module.scss';
import { ISpfxLoaderProps } from './ISpfxLoaderProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

export interface ISpfxLoaderState {
  loading: boolean;
  items: any[];
}

export default class SpfxLoader extends React.Component<ISpfxLoaderProps, ISpfxLoaderState> {

  constructor(props: ISpfxLoaderProps) {
    super(props);
    this.state = {
      loading: false,
      items: []
    };
  }

  public componentDidMount() {
    this.getItems();
  }

  public getItems(){
    let graphURI: string = "/sites/root/lists";

    if (!this.props.graphClient) {
      return;
    }
    this.setState({
      loading: true,
    });

    this.props.graphClient
      .api(graphURI)
      .version("v1.0")
      .get((err: any, res: any): void => {
        if (err) {
          this.setState({
            loading: false
          });
          return;
        }
        if (res && res.value && res.value.length > 0) {
          console.log("res: ", res);
          this.setState({
            items: res.value,
            loading: false
          });
        }
        else {
          this.setState({
            loading: false
          });
        }
      });

  }

  public render(): React.ReactElement<ISpfxLoaderProps> {
    return (
      <div className={styles.spfxLoader}>
        <h2>{this.props.description}</h2>
        {
          this.state.loading &&
          <Spinner label="Loading items..." size={SpinnerSize.large} />
        }
        {
          this.state.items.length > 0 && this.state.items.map(m => <p>{m.name}</p>)
        }
      </div>
    );
  }
}
