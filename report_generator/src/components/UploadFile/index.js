import React from "react"
import Dropzone from 'react-dropzone'
import './style.css';
import HelperServices from "../../services"
import { Dropdown, Row, Form, Col } from 'react-bootstrap'
import CircularProgress from '@material-ui/core/CircularProgress';

import Button from '@material-ui/core/Button';

class UploadFile extends React.Component {
    constructor(props) {
        super(props)
        this.state = {
            testEngineerNameList: ["Select a name", "Ankit Kumar", "Kaushal Kumar", "Mohit Singh", "Jatin", "Avishek Kumar", "Parth", "Tushant Rajvanshi", "Isha Sachdev", "Sumit Saklani", "Aviral mishra", "Tripti Tiwari", "Kajal Jha","Gaurav Goswami"],
            selectedEngineerName: "Select a name",
            approvingAuthority: "Zahid Raza",
            approvingAuthorityList: ["Zahid Raza", "Shashank Raghubanshi"],
            file: null,
            fileName: "",
            rows: null,
            showLoader: false
        }
    }
    componentDidMount() {
        HelperServices.getReportFileList()
            .then((data) => {
                this.setState({
                    rows: data
                })
            })
    }
    onFileUpload = (files) => {
        if (files.length !== 0) {
            let file = files[0]
            let fileName = file.name
            this.setState({
                file,
                fileName
            })
        }
    }

    submitReport = () => {
        let { selectedEngineerName,approvingAuthority, file, fileName, rows } = this.state

        if (selectedEngineerName == "Select a name") {
            alert("select a engineer name")
        }
        else {
            if (file) {
                let name = fileName.split('.').slice(0, -1).join('.')
                if (!rows.find(file => file.ReportName == name)) {
                    this.setState({
                        showLoader: true
                    })
                    HelperServices.uploadFile(selectedEngineerName, file, fileName, approvingAuthority)
                        .then((res) => {
                            this.setState({
                                showLoader: false
                            })
                        })
                        .catch((e) => {
                            this.setState({
                                showLoader: false
                            })
                            alert("Upload failed")
                        })

                }
                else {
                    alert("The file name is duplicate, you cant upload a file with same name till its rejected")
                }
            }
            else {
                alert("select a Report File")
            }
        }

    }

    onNameChnage = (name) => {
        this.setState({
            selectedEngineerName: name
        })
    }
    onApprovingAthChnage = (name) => {
        this.setState({
            approvingAuthority: name
        })
    }

    render() {
        return (
            <div style={{ position: "relative" }}>
                {
                    this.state.showLoader ? <div className="loader-class">
                        <CircularProgress className="loader" disableShrink />
                    </div> : null
                }
                <div style={{ opacity: this.state.showLoader ? 0.5 : 1 }}>
                    <div className="enginer-name-div text-align-left" >
                        <span className="display-block pr-2">Select Tese Engineer Name</span>
                        <Dropdown variant="contained" color="primary" className="display-block pl-4 ">
                            <Dropdown.Toggle variant="success" id="dropdown-basic">
                                {this.state.selectedEngineerName}
                            </Dropdown.Toggle>

                            <Dropdown.Menu className="btn-blue" >
                                {
                                    this.state.testEngineerNameList.map(name =>
                                        <Dropdown.Item onClick={() => this.onNameChnage(name)}>{name}</Dropdown.Item>
                                    )
                                }
                            </Dropdown.Menu>
                        </Dropdown>
                        <span className="display-block pr-2" style={{marginLeft:"15px"}}>Technical Manager</span>
                        <Dropdown variant="contained" color="primary" className="display-block pl-4 ">
                            <Dropdown.Toggle variant="success" id="dropdown-basic">
                                {this.state.approvingAuthority}
                            </Dropdown.Toggle>

                            <Dropdown.Menu className="btn-blue" >
                                {
                                    this.state.approvingAuthorityList.map(name =>
                                        <Dropdown.Item onClick={() => this.onApprovingAthChnage(name)}>{name}</Dropdown.Item>
                                    )
                                }
                            </Dropdown.Menu>
                        </Dropdown>
                   
                    </div>
                    <Dropzone onDrop={this.onFileUpload} accept={[".docx"]} >
                        {({ getRootProps, getInputProps }) => (
                            <div className="container-fluid p-3">
                                <section className="card-border">
                                    <div {...getRootProps()} className="drop-div" >
                                        <div className="move-to-tenter" >
                                            <div className="container-fluid">
                                                <div className="row pt-3" >
                                                    <div className="col-12 ">
                                                        <input {...getInputProps()} />
                                                        <img style={{ height: "150px" }} src={require('../../assets/images/cloud-upload.svg')} alt="cloud-icon" />
                                                    </div>
                                                </div>
                                                <div className="row pt-2" >
                                                    <div className="col-12">
                                                        <span className="new-template source-sans-pro" style={{ fontWeight: "500", fontSize: "x-large" }}>Drag File here or</span>

                                                    </div>
                                                </div>
                                                <div className="row pt-2" >
                                                    <div className="col-12">
                                                        <button type="button" className="btn source-sans-pro button-text-class browse-image-btn" onChange={this.onFileUpload}>Browse for file</button>

                                                    </div>
                                                </div>
                                                <div className="row pt-4" >
                                                    <div className="col-12">
                                                        <span className="source-sans-pro" style={{ fontSize: "13px", marginTop: "10px" }}>Drop your file here</span>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </section>
                            </div>
                        )}
                    </Dropzone>
                    <Form>
                        <Form.Group controlId="reportFileName">
                            <Row>
                                <Form.Label className="text-align-left" column sm={2}>Report File Name</Form.Label>
                                <Col sm={4}>
                                    <Form.Control placeholder="File Name" readonly value={this.state.fileName} />
                                </Col>
                            </Row>
                            <Row className="pt-4" style={{ display: "block" }}>
                                <div style={{ float: "right" }}>
                                    <Button variant="contained" color="primary" onClick={this.submitReport}>Submit Report</Button>
                                </div>
                            </Row>

                        </Form.Group>
                    </Form>
                </div>

            </div>
        )
    }
}

export default UploadFile;