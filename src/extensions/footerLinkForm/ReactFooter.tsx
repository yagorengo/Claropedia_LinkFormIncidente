import * as React from "react";  
import { CommandButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';  
import { IContextualMenuProps, IIconProps } from 'office-ui-fabric-react';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { sp } from "@pnp/sp/presets/all";
import styles from "./ReactFooter.module.scss";


export interface IReactFooterProps {
    context: ApplicationCustomizerContext;
}  
  export interface IREactFooterState {
      items:any
  }
export default class ReactFooter extends React.Component<IReactFooterProps, IREactFooterState> {  
  constructor(props: IReactFooterProps) {  
    super(props);  
    
  }  
  protected async onInit(): Promise<void> {
    sp.setup(this.props.context);
  }
  
  private iconProps:IIconProps = {
    iconName:'OpenEnrollment', 
    styles: {root:{color:'white'}}
  }
  private iconPropsConsultas:IIconProps = {
    iconName:'Questionnaire', 
    styles: {root:{color:'white'}}
  }

  public render(): JSX.Element {  
   //let href =`https://claroaup.sharepoint.com/sites/BackOfficeClaropedia/Lists/test%20Reportes%20de%20Incidentes/NewForm.aspx?`
   let hrefMisConsultas = "https://claroaup.sharepoint.com/sites/BackOfficeClaropedia/Lists/test%20Reportes%20de%20Incidentes/AllItems.aspx"
    return (  
      <div className={styles.app} >  
         <CommandButton className={styles.button} data-interception="off" iconProps={this.iconProps} target="_blank" text="Contactanos" onClick={()=>this.onClickContact()}/>
         <CommandButton className={styles.button} data-interception="off" iconProps={this.iconPropsConsultas} target="_blank" text="Mis consultas" href={hrefMisConsultas}/>
      </div>  
    );  
  }  

  private onClickContact = ():void => {
    let page = window.location.pathname;
    let hostname = window.location.host;
    let fullUrl = 'https://'+hostname+page;
    let pais = this.props.context.pageContext.web.title.substring(11)
    //console.log("page", this.props.context.pageContext.listItem.id);
    sp.web.lists.getByTitle("Páginas del sitio").items.getById(this.props.context.pageContext.listItem.id).get().then((res)=> {
      window.open(`https://claroaup.sharepoint.com/sites/BackOfficeClaropedia/SitePages/Formulario%20de%20incidentes.aspx?link=${fullUrl}&nombre=${res.Title}&pais=${pais}`, '_blank');
    })
  }
  
}  