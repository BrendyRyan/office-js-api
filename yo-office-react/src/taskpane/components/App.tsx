import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default function App(appProps:AppProps){
const [appState, setAppState] = React.useState([])

React.useEffect(()=>{
  setAppState([ {
    icon: "Ribbon",
    primaryText: "Achieve more with Office integration",
  },
  {
    icon: "Unlock",
    primaryText: "Unlock features and functionality",
  },
  {
    icon: "Design",
    primaryText: "Create and visualize like a pro",
  },
  {
    icon: "OfficeLogo",
    primaryText: "A test for the add in",
  },])
},[])

async function handleClick(){  
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
};

const { title, isOfficeInitialized } = appProps;

if (!isOfficeInitialized) {
  return (
    <Progress
      title={title}
      logo={require("./../../../assets/logo-filled.png")}
      message="Please sideload your addin to see app body."
    />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={appProps.title} message="Welcome" />
      <HeroList message="Discover what Office Add-ins can do for you today!" items={appState}>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
          Run
        </DefaultButton>
      </HeroList>
    </div>
  );
}