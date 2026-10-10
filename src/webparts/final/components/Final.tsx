import * as React from "react";
import type { IFinalProps } from "./IFinalProps";
import { get_SP } from "../../../PnpConfig";
import { IFinalStates, IStates } from "../../../IFinalStates";
import styles from "./Final.module.scss";

import {
  DatePicker,
  Dropdown,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  TextField,
  Toggle,
} from "@fluentui/react";

export default class Final extends React.Component<IFinalProps, IFinalStates> {
  constructor(props: IFinalProps) {
    super(props);

    this.state = {
      emps: [],
      Title: "",
      Age: 0,
      isActive: false,
      Dept: "",
      DeptOptions: [],
      Date: undefined,
      Sal: 0,
      showForm: false,
      submitMsg: "",
      emptyData: "",
      dltMsg: "",
      newEdit: "",
      editId: "",
      searchText: "",
     
    };
  }

  private sp = get_SP();
  public get_SPItems = async (): Promise<void> => {
    try {
      const res = await this.sp.web.lists
        .getByTitle("EmployessList")
        .items.select(
          "ID",
          "Title",
          "Age",
          "Dept",
          "Sal",
          "isActive",
          "Date",
        )();

      this.setState({ emps: res });

      console.log(res, "res");
    } catch (error) {
      console.log(error);
    }
  };

  async componentDidMount(): Promise<void> {
    await this.get_SPItems();
    await this.checkOptions();
  }
  public newEmp = (): void => {
    this.setState({
      newEdit: "Save",
      showForm: true,
    });
  };

  public formatDate = (date: Date) => {
    const year = date.getFullYear();
    const month = ("0" + (date.getMonth() + 1)).slice(-2);
    const day = ("0" + date.getDate()).slice(-2);

    return `${year}-${month}-${day}`;
  };

  public createForm = async (): Promise<void> => {
    if (!this.state.Title || !this.state.Age || !this.state.Date) {
      this.setState({
        emptyData: "Please enter data before submit,fields are required",
      });

      setTimeout(() => {
        this.setState({ emptyData: "" });
      }, 3000);

      return;
    } else {
      try {
        await this.sp.web.lists.getByTitle("EmployessList").items.add({
          Title: this.state.Title,
          Age: this.state.Age,
          Dept: this.state.Dept,
          Sal: this.state.Sal,
          isActive: this.state.isActive,
          Date: this.state.Date ? this.formatDate(this.state.Date) : undefined,
        });

        this.setState({
          Title: "",
          Age: 0,
          Dept: "",
          Sal: 0,
          isActive: false,
          Date: undefined,
          submitMsg: "Data submitted successfully",
          showForm: false,
        });

        await this.get_SPItems();

        setTimeout(() => {
          this.setState({ submitMsg: "" });
        }, 3000);
      } catch (error) {
        console.log(error);
      }
    }
  };
  checkOptions = async (): Promise<void> => {
    const options = await this.sp.web.lists
      .getByTitle("EmployessList")
      .fields.getByInternalNameOrTitle("Dept")();

    this.setState({
      DeptOptions: options.Choices,
    });
  };

  public CancelForm = (): void => {
    this.setState({
      showForm: false,
      Title: "",
      Age: 0,
      Dept: "",
      Sal: 0,
      isActive: false,
      Date: undefined,
    });
  };
  public CancelData = (): void => {
    this.setState({
      Title: "",
      Age: 0,
      Dept: "",
      Sal: 0,
      isActive: false,
      Date: undefined,
    });
  };

  public handleDelete = async (dltID: any): Promise<void> => {
    try {
      await this.sp.web.lists
        .getByTitle("EmployessList")
        .items.getById(dltID)
        .delete();

      this.setState({
        dltMsg: "Item deleted successfully",
      });

      await this.get_SPItems();

      setTimeout(() => {
        this.setState({ dltMsg: "" });
      }, 3000);
    } catch (error) {
      console.log(error);
    }
  };
  public handleEdit = (item: any) => {
    this.setState({
      newEdit: "update",
      showForm: true,
      editId: item.ID,
      Title: item.Title,
      Age: item.Age,
      Dept: item.Dept,
      Sal: item.Sal,
      isActive: item.isActive,
      Date: item.Date ? new Date(item.Date) : undefined,
    });
  };

  public editForm = async (): Promise<void> => {
    try {
      await this.sp.web.lists
        .getByTitle("EmployessList")
        .items.getById(Number(this.state.editId))
        .update({
          Title: this.state.Title,
          Age: this.state.Age,
          Dept: this.state.Dept,
          Sal: this.state.Sal,
          isActive: this.state.isActive,
          Date: this.state.Date ? this.formatDate(this.state.Date) : undefined,
        });

      this.setState({
        Title: "",
        Age: 0,
        Dept: "",
        Sal: 0,
        isActive: false,
        Date: undefined,
        submitMsg: "Item updated successfully",
        newEdit: "",
        editId: "",
        showForm: false,
      });

      await this.get_SPItems();

      setTimeout(() => {
        this.setState({ submitMsg: "" });
      }, 3000);
    } catch (error) {
      console.error("Update error:", error);
    }
  };

  public render(): React.ReactElement<IFinalProps> {
    return (
      <>
        {this.state.showForm && (
          <div className={styles.overlay}>
            <div className={styles.newForm}>
              <li onClick={this.CancelForm} className={styles.li}>
                X
              </li>

              {this.state.emptyData && (
                <MessageBar messageBarType={MessageBarType.warning}>
                  {this.state.emptyData}
                </MessageBar>
              )}

              <TextField
                label="Enter your Name"
                value={this.state.Title}
                onChange={(e, v) => this.setState({ Title: v || "" })}
              />

              <TextField
                label="Enter your Age"
                value={this.state.Age.toString()}
                onChange={(e, v) => this.setState({ Age: Number(v) })}
              />

              <Dropdown
                options={(this.state.DeptOptions || []).map((opt) => ({
                  key: opt,
                  text: opt,
                }))}
                label="select your skills"
                placeholder="select your skills"
                selectedKey={this.state.Dept}
                onChange={(_, option) =>
                  this.setState({
                    Dept: option?.key as string,
                  })
                }
              />

              <TextField
                label="Enter your Sal"
                value={this.state.Sal.toString()}
                onChange={(_, v) => this.setState({ Sal: Number(v) })}
              />

              <DatePicker
                label="DateofJoining"
                value={this.state.Date}
                onSelectDate={(date) =>
                  this.setState({
                    Date: date || undefined,
                  })
                }
              />

              <Toggle
                label="Active?"
                onText="Yes"
                offText="No"
                checked={this.state.isActive}
                onChange={(e, v) =>
                  this.setState({
                    isActive: v || false,
                  })
                }
              />

              <PrimaryButton
                text={this.state.newEdit}
                style={{
                  color: "white",
                  marginRight: "10px",
                  backgroundColor: "green",
                  marginLeft: "80px",
                }}
                onClick={
                  this.state.newEdit === "Save"
                    ? this.createForm
                    : this.editForm
                }
              />

              <PrimaryButton
                text="cancel"
                style={{
                  color: "white",
                  backgroundColor: "red",
                  marginLeft: "110px",
                }}
                onClick={this.CancelData}
              />
            </div>
          </div>
        )}

        {!this.state.showForm && (
          <button onClick={this.newEmp} className={styles.btnEmp}>
            Create
          </button>
        )}

        {this.state.dltMsg && (
          <MessageBar messageBarType={MessageBarType.success}>
            {this.state.dltMsg}
          </MessageBar>
        )}

        {this.state.submitMsg && (
          <MessageBar messageBarType={MessageBarType.success}>
            {this.state.submitMsg}
          </MessageBar>
        )}

        <TextField
          placeholder="Search by Employee Name..."
          value={this.state.searchText}
          onChange={(_, value) =>
            this.setState({
              searchText: value || "",
            })
          }
        />
        <table className={styles.table}>
          <thead>
            <tr>
              <th>EmpID</th>
              <th>EmpName</th>
              <th>EmpAge</th>
              <th>EmpDept</th>
              <th>EmpSal</th>
              <th>HiredDate</th>
              <th>isActive</th>
              <th>Edit</th>
              <th>Delete</th>
            </tr>
          </thead>

          <tbody>
            {this.state.emps.map((items: IStates) => {
              return (
                <tr key={items.ID}>
                  <td>{items.ID}</td>
                  <td>{items.Title}</td>
                  <td>{items.Age}</td>
                  <td>{items.Dept}</td>
                  <td>{items.Sal}</td>
                  <td>{items.Date}</td>
                  <td>{items.isActive ? "yes" : "no"}</td>

                  <td
                    className={styles.modify}
                    onClick={() => this.handleEdit(items)}
                  >
                    ✏️
                  </td>

                  <td
                    className={styles.modify}
                    onClick={() => this.handleDelete(items.ID)}
                  >
                    🗑️
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </>
    );
  }
}
