import React, { useState, useEffect } from 'react';
import { DetailsList, SelectionMode, DefaultButton } from '@fluentui/react';
import { app, call, mail } from '@microsoft/teams-js'
import { Button, Card, CardPreview, CardHeader, Text, Caption1, makeStyles, shorthands, tokens, Checkbox } from "@fluentui/react-components";
import {
  CallRegular,
  CalendarMailRegular,
  MoreHorizontal20Filled 
} from "@fluentui/react-icons";

export const SuppliersCard = () => {
  const [suppliers, setSuppliers] = useState([]);
  const [selectedSupplier, setSelectedSupplier] = useState(null);

  useEffect(() => {
    async function fetchSuppliers() {
      const response = await fetch("https://services.odata.org/V4/Northwind/Northwind.svc/Suppliers");
      const data = await response.json();
      setSuppliers(data.value);

      try {
        const searchParams = window.location.href;
        await app.initialize();
        const context = await app.getContext();
        if (searchParams.includes('country')) {
          const country = searchParams.match(/=(.*)/)[1];
          const supplier = data.value.filter(x => x.Country === country)
          if (supplier.length > 0) {
            setSuppliers(supplier);
          }
        }
        //deeplinking
        if (context.page.subPageId) {
          const supplier = data.value.filter(x => x.SupplierID === context.page.subPageId)
          if (supplier.length > 0) {
            setSelectedSupplier(supplier[0]);
          }
          else {
            console.error('Supplier not found or invalid data');
          }
        }
      } catch (error) {
        console.error('Could not initialize Teams JS client library');
      }
    }
    fetchSuppliers();
  }, []);

  const handleRowClick = (supplier) => {
    setSelectedSupplier(supplier);
  };
  const styles = useStyles();

  const memoCardList = React.useMemo( () =>    
  (<CardListExample 
    selected={selectedSupplier}
    onselectedChange={(_, { selected }) => handleRowClick(selected)}
    suppliers={suppliers}
    call={call}
    mail={mail}
  />), [selectedSupplier, suppliers]);

  return (
    <div className={styles.main}>
      {/* <CardExample
        selected={selectedSupplier}
        onSelectionChange={(_, { selected }) => handleRowClick(selected)}
        supplierName={"What"}
      />
      <CardExample
        selected={selectedSupplier}
        onSelectionChange={(_, { selected }) => handleRowClick(selected)}
        supplierName={"What"}
      /> */}
      {memoCardList}
    </div>
  );
};

const renderContactButton = (item, call, mail) => {
  if (call.isSupported()) {
    return (
      <Button
        appearance="transparent"
        icon={<CallRegular />}
        onClick={async () => {
          await call.startCall({
            targets: [
              'adeleV@m365404404.onmicrosoft.com',
              'admin@m365404404.onmicrosoft.com'
            ],
            requestedModalities: [
              call.CallModalities.Audio,
              call.CallModalities.Video,
              call.CallModalities.VideoBasedScreenSharing,
              call.CallModalities.Data
            ]
          });
        }}
      ></Button>
    );
  } else if (mail.isSupported()) {
    return (
      <Button
        appearance="transparent"
        icon={<CalendarMailRegular />}
        onClick={async () => {
          mail.composeMail({
            type: mail.ComposeMailType.New,
            subject: `Enquiry for supplier:${item.CompanyName}`,
            message: 'Hello',
            toRecipients: [
              'adeleV@m365404404.onmicrosoft.com',
              'admin@m365404404.onmicrosoft.com'
            ]
          });
        }}
      ></Button>
    );
  }
};

const useStyles = makeStyles({
  main: {
    ...shorthands.gap("16px"),
    display: "flex",
    flexWrap: "wrap",
  },

  row: {
    ...shorthands.gap("16px"),
    display: "flex",
    flexWrap: "wrap",
  },

  card: {
    width: "270px",
    maxWidth: "100%",
    height: "fit-content",
  },

  caption: {
    color: tokens.colorNeutralForeground3,
  },

  smallRadius: {
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
  },

  grayBackground: {
    backgroundColor: tokens.colorNeutralBackground3,
  },

  logoBadge: {
    ...shorthands.padding("5px"),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    backgroundColor: "#FFF",
    boxShadow:
      "0px 1px 2px rgba(0, 0, 0, 0.14), 0px 0px 2px rgba(0, 0, 0, 0.12)",
  },

  logo: {
    ...shorthands.borderRadius("4px"),
    width: "48px",
    height: "48px",
  },

  actions: {
    display: "flex",
  },
});

const CardExample = (props) => {
  const styles = useStyles();

  return (
    <Card className={styles.card} {...props}>
      <CardPreview
        className={styles.grayBackground}
        logo={
          <img
            className={styles.logoBadge}
            src={"https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/assets/logo3.svg"}
            alt="Figma app logo"
          />
        }
      >
        <img
          className={styles.smallRadius}
          src={"https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/assets/office1.png"}
          alt="Presentation Preview"
        />
      </CardPreview>

      <CardHeader
        header={<Text weight="semibold">{props.supplierName}</Text>}
        description={
          <Caption1 className={styles.caption}>You created 53m ago</Caption1>
        }
        action={
          <Button
            appearance="transparent"
            icon={<MoreHorizontal20Filled />}
            aria-label="More actions"
          />
        }
      />
    </Card>
  );
};

const CardListExample = (props) => {
  const styles = useStyles();
  const imageSrc = [
    "https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/assets/logo.svg",
    "https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/assets/logo2.svg",
    "https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/assets/app_logo.svg",
    "https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/assets/office1.png"   
  ];

  return (
    <div className={styles.row}>
      <h2>Suppliers</h2>
      {props.suppliers.map((_item, _index) => {
          const img = imageSrc[Math.floor(Math.random() * props.suppliers.length)];
          return (
          <Card
          className={styles.card}
          selected={props.selected}
          onSelectionChange={props.onselectedChange}
          floatingAction={
            // <Checkbox onChange={props.onselectedChange} checked={props.selected} />
            renderContactButton(_item, props.call, props.mail)
          }
        >
          <CardHeader
            image={
              <img
                className={styles.logo}
                src={img}
                alt="Logo"
              />
            }
            header={<Text weight="semibold">{_item.CompanyName}</Text>}
            description={
              <Caption1 className={styles.caption}>
                {_item.ContactName}
              </Caption1>
            }
          />
        </Card>)
        })
      }
      {/* <Card
          className={styles.card}
          selected={props.selected}
          onSelectionChange={props.onselectedChange}
          floatingAction={
            <Checkbox onChange={props.onselectedChange} checked={props.selected} />
          }
        >
          <CardHeader
            image={
              <img
                src={"https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/assets/word_logo.svg"}
                alt="Microsoft Word Logo"
              />
            }
            header={<Text weight="semibold">Secret Project Briefing</Text>}
            description={
              <Caption1 className={styles.caption}>
                OneDrive &gt; Documents
              </Caption1>
            }
          />
        </Card>

        <Card
          className={styles.card}
          selected={props.selected}
          onSelectionChange={props.onselectedChange}
          floatingAction={
            <Checkbox onChange={props.onselectedChange} checked={props.selected} />
          }
        >
          <CardHeader
            image={
              <img
                src={"https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/assets/excel_logo.svg"}
                alt="Microsoft Excel Logo"
              />
            }
            header={<Text weight="semibold">Team Budget</Text>}
            description={
              <Caption1 className={styles.caption}>
                OneDrive &gt; Spreadsheets
              </Caption1>
            }
          />
        </Card> */}
      </div>
  )
}
