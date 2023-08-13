import * as React from 'react';
import { useState } from 'react';
import { ICrudOperationsProps } from './ICrudOperationsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.min.js';
interface EmployeeStates {
  ID: number;
  Title: string;
  Age: string;
}

const CrudOperations: React.FC<ICrudOperationsProps> = (props: ICrudOperationsProps) => {
  const [fullName, setFullName] = useState('');
  const [age, setAge] = useState('');
  const [allItems, setAllItems] = useState<EmployeeStates[]>([]);

  //Create Item
  const createItem = async (): Promise<void> => {
    const body: string = JSON.stringify({
      'Title': fullName,
      'Age': age
    });
    try {
      const response: SPHttpClientResponse = await props.context.spHttpClient.post(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeData')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: body
        }
      );
      if (response.ok) {
        const responseJSON = await response.json();
        console.log(responseJSON);
        alert(`Item created successfully with ID: ${responseJSON.ID}`);
      }
      else {
        const responseJSON = await response.json();
        console.log(responseJSON);
        alert(`Something went wrong! please check the browser window for issues`);
      }
    }
    catch (err) {
      console.log(err);
      alert(`An occurred while creating item`);
    }
  }
  //Get Item BY ID
  const getItemByID = (): void => {
    const idElement = document.getElementById('itemId') as HTMLInputElement | null;
    if (idElement?.value) {
      const id: number = Number(idElement.value); // Make sure to convert the value to a number
      if (id > 0) {
        props.context.spHttpClient.get(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeData')/items(${id}) `,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }
        )
          .then((response: SPHttpClientResponse) => {
            if (response.ok) {
              response.json().then((responseJSON) => {
                setFullName(responseJSON.Title);
                setAge(responseJSON.Age);
              });

            }
            else {
              response.json().then((responseJSON) => {
                console.log(responseJSON);
                alert(`Something went wrong! please check the browser window for the issues`);
              });
            }
          })
          .catch((err) => {
            console.log(err);
          });
      }
      else {
        alert(`Please enter the valid id`);
      }
    }
    else {
      console.log("Error: Element 'itemID' not found. ");
    }
  }
  //GetAlitems

  const getAllItems = (): void => {
    props.context.spHttpClient.get(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeData')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }
    )
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((responseJSON) => {
            setAllItems(responseJSON.value);
            console.log(responseJSON);
          });
        }
        else {
          response.json().then((responseJSON) => {
            console.log(responseJSON);
            alert(`Something went wrong! please check the browser console for more information`);
          });
        }
      })
      .catch((err) => {
        console.log(err);

      });
  }
  //Update Data
  const updateItem = (): void => {
    const idElement = document.getElementById('itemId') as HTMLInputElement | null;
    if (idElement) {
      const id: number = parseInt(idElement.value);
      const body: string = JSON.stringify({
        'Title': fullName,
        'Age': parseInt(age),
      });
      if (id > 0) {
        props.context.spHttpClient.post(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeData')/items(${id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': '',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE'
            },
            body: body
          }
        )
          .then((response: SPHttpClientResponse) => {
            if (response.ok) {
              alert(`Item with ID: ${id} updated successfully `);
            }
            else {
              response.json().then((responseJSON) => {
                console.log(responseJSON);
                alert(`Something went wrong ! check the error in the browser console`);
              });
            }
          })
          .catch((err) => {
            console.log(err);
          })
      }
      else {
        alert(`Please enter the valid item id. `);
      }
    }
    else {
      alert(`Item ID element is not found. `);
    }
  };

  // Delete Item
  const deleteItem = (): void => {
    const idElement = document.getElementById('itemId') as HTMLInputElement 
    const id: number = parseInt(idElement?.value || '0');
    if (id > 0) {
      props.context.spHttpClient.post(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeData')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }
      )
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            alert(`Item ID : ${id} deleted successfully`);
          }
          else {
            alert(`something went wrong! check the error in browser console`);
            console.log(response.json());
          }
        });
    }
    else {
      alert(`Please eneter a valid item id. `);
    }
  }
  return (
    <>
      <div className="container">
        <div className="row">
          <div className="col-md-6">
            <p>{escape(props.description)}</p>
            <div className="form-group">
              <label htmlFor="itemId">Item ID:</label>
              <input type="text" className="form-control" id="itemId"></input>
            </div>
            <div className="form-group">
              <label htmlFor="fullName">Full Name</label>
              <input type="text" className="form-control" id="fullName" value={fullName} onChange={(e) => setFullName(e.target.value)}></input>
            </div>
            <div className="form-group">
              <label htmlFor="age">Age</label>
              <input type="text" className="form-control" id="age" value={age} onChange={(e) => setAge(e.target.value)}></input>
            </div>
            <div className="form-group">
              <label htmlFor="allItems">All Items:</label>
              <div id="allItems">
                <table className="table table-bordered">
                  <thead>
                    <tr>
                      <th>ID</th>
                      <th>Full Name</th>
                      <th>Age</th>
                    </tr>
                  </thead>
                  <tbody>
                    {allItems.map((item) => (
                      <tr key={item.ID}>
                        <td>{item.ID}</td>
                        <td>{item.Title}</td>
                        <td>{item.Age}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div className="d-flex justify-content-start">
              <button className="btn btn-primary mx-2" onClick={createItem}>Create</button>
              <button className="btn btn-success mx-2" onClick={getItemByID}>Read</button>
              <button className="btn btn-info mx-2" onClick={getAllItems}>Read All</button>
              <button className="btn btn-warning mx-2" onClick={updateItem}>Update</button>
              <button className="btn btn-danger mx-2" onClick={deleteItem}>Delete</button>
            </div>
          </div>
        </div>
      </div>

    </>
  )
}
export default CrudOperations