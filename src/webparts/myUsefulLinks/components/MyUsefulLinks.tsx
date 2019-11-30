import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from './MyUsefulLinks.module.scss';
import { IMyUsefulLinksProps } from './IMyUsefulLinksProps';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Panel } from 'office-ui-fabric-react/lib/Panel';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { IMyUsefulLinksState } from './IMyUsefulLinksState';
import { sp, ItemAddResult } from "@pnp/sp";

export default class MyUsefulLinks extends React.Component<IMyUsefulLinksProps, IMyUsefulLinksState> {

  public constructor(props: IMyUsefulLinksProps, state: IMyUsefulLinksState) {
    super(props);
    this.state = { 
      isOpen: false, 
      PanelTitle: '',
      title: '',
      link: '',
      id: 0 
    };
  }

  public openPane():void {
    this.setState({ isOpen: true }); 
  }

  public dismissPanel():void {
    this.setState({ isOpen: false });
  }
  
  public items:any;
  public defaultItems:any;

  public AddLinkButton = () => {
    this.setState({ PanelTitle: 'Add link', title: '', link:'', id:0 });
    this.openPane();
  }
  public AddLink = () => {
    if(this.state.title == '' || this.state.link == ''){
      alert('All fields are required');
    } else {
      ReactDom.render( <div className={ styles.usefulLinks }><Spinner className={styles.spinner} label="Please wait..." size={SpinnerSize.large} ariaLive="assertive" labelPosition="right" /></div>, this.props.tsThis.domElement);
      this.dismissPanel();
      //console.info(this.state.title, this.state.link, this.state.id);
      if(this.state.id == 0){
        sp.web.lists.getByTitle('My Useful Links').items.add({
          Title: this.state.title,
          Url: this.state.link
        }).then((iar: ItemAddResult) => {
          console.info(iar);
          this.props.render(this.props.tsThis);
        });
      } else {
        sp.web.lists.getByTitle('My Useful Links').items.getById(this.state.id).update({
          Title: this.state.title,
          Url: this.state.link
        }).then(i => {
          console.info(i);
          this.props.render(this.props.tsThis);
        });
      }
    }
    
  }
  public EditLinkButton = (item) => {
    this.setState({ PanelTitle: 'Edit link', title: item.Title, link:item.Url, id:item.Id });
    this.openPane();
  }
  public DeleteLinkButton = (item) => {
    let del = confirm('Do you want to delete this link?');
    if(del){
      console.info('Delete id:'+item.Id);
      ReactDom.render( <div className={ styles.usefulLinks }><Spinner className={styles.spinner} label="Please wait..." size={SpinnerSize.large} ariaLive="assertive" labelPosition="right" /></div>, this.props.tsThis.domElement);
      sp.web.lists.getByTitle('My Useful Links').items.getById(item.Id).recycle().then(_ => {this.props.render(this.props.tsThis);});
    }
  }

  public titleChange(event):void  {
    this.setState({ title: event.target.value });
  }

  public urlChange(event):void {
    this.setState({ link: event.target.value });
  }

  public render(): React.ReactElement<IMyUsefulLinksProps> {

    //console.info(this.props.defaultLinks);
    this.defaultItems = this.props.defaultLinks.map((item, key) => {
      let target = item.OpenNewWindow ? '_blank' : 'self';
      return  <div className={styles.userfulLinksItem}>
                <a href={item.Url} target={target} className={styles.userfulLink}>{item.Title}</a>
              </div>;
    });
    
    this.items = this.props.myLinks.map((item, key) => {
      let target = item.OpenNewWindow ? '_blank' : 'self';
      return  <div className={styles.userfulLinksItem}>
                <a href={item.Url} target={target} className={styles.userfulLink}>{item.Title}</a>
                <div className={styles.userfulLinkButtons}>
                  <Icon iconName='Edit' className={styles.IconButton} title="Edit" onClick={() => this.EditLinkButton(item)} />
                  <span> </span>
                  <Icon iconName='Delete' className={styles.IconButton} title="Delete" onClick={() => this.DeleteLinkButton(item)} />
                </div>
              </div>;
    });

    return (
      <div className={ styles.usefulLinks }>
        <div className={styles.wptitle}>
          <Icon iconName='Link' className={styles.wptitleIcon} />
          <span>{this.props.title}</span>
          <div className={styles.addEventContainer} title="Add link"  onClick={this.AddLinkButton}>
            <Icon iconName='CircleAddition' className={styles.wptitleIcon} />
          </div>
        </div>
        <div className={styles.userfulLinksItems}>
          {this.defaultItems}
          {this.items}
        </div>
        <Panel
          headerText={this.state.PanelTitle}
          isOpen={this.state.isOpen}
          onDismiss={() => this.dismissPanel()}
          closeButtonAriaLabel="Close"
          className={styles.panel}
        >
            <TextField label="Title" onChange={(event) => this.titleChange(event)} value={this.state.title} required />
            <TextField label="URL" onChange={(event) => this.urlChange(event)} value={this.state.link} required />
            <div className={styles.panelToolbar}>
              <PrimaryButton text="Save" onClick={this.AddLink} allowDisabledFocus />
              <span> </span>
              <DefaultButton text="Cancel" onClick={() => this.dismissPanel()} allowDisabledFocus />
            </div>
        </Panel>
      </div>
    );
  }
}
