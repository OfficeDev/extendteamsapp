import { DefaultButton, DetailsList, SelectionMode } from '@fluentui/react';
import { Link, Text } from '@fluentui/react-components'
import React, { useEffect, useState } from 'react';

import { app } from '@microsoft/teams-js'

export const Suppliers = (props) => {
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

  const supplierColumns = [
    {
      key: 'companyName',
      name: 'Name',
      fieldName: 'CompanyName',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
      onRender: (item) => {
        return (
          <Link key={item.id} style={{ fontSize: '12px' }} onClick={() => handleRowClick(item)}>
            {item.CompanyName}
          </Link>
        );
      }
    },
    {
      key: 'contactName',
      name: 'Contact',
      fieldName: 'ContactName',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'phone',
      name: 'Phone',
      fieldName: 'Phone',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'country',
      name: 'Country',
      fieldName: 'Country',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    }
  ];
  const selectSuppliercolumn = [
    {
      key: 'companyName',
      name: 'Name',
      fieldName: 'CompanyName',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'contactName',
      name: 'Contact',
      fieldName: 'ContactName',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'phone',
      name: 'Phone',
      fieldName: 'Phone',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'country',
      name: 'Country',
      fieldName: 'Country',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    }
  ];

  return (
    <div>
      {!selectedSupplier && (
        <div>
          <div className='headingSupplier'>
            <Text size={500} as="h2" style={{ margin: "15px" }}>Suppliers</Text>
          </div>
          <DetailsList
            items={suppliers}
            columns={supplierColumns}
            selectionMode={SelectionMode.single}
            onItemInvoked={handleRowClick}
          />
        </div>)}
      {selectedSupplier && (
        <div>
          <DetailsList
            items={[selectedSupplier]}
            columns={selectSuppliercolumn}
            selectionMode={SelectionMode.none}
          />
          <DefaultButton key={""} onClick={() => handleRowClick(null)}>
            Back to Suppliers
          </DefaultButton>
        </div>
      )}
    </div>
  );
};
