import React from 'react';
import { makeStyles } from '@material-ui/core/styles';
import Paper from '@material-ui/core/Paper';
import Table from '@material-ui/core/Table';
import TableBody from '@material-ui/core/TableBody';
import TableCell from '@material-ui/core/TableCell';
import TableContainer from '@material-ui/core/TableContainer';
import TableHead from '@material-ui/core/TableHead';
import TablePagination from '@material-ui/core/TablePagination';
import TableRow from '@material-ui/core/TableRow';
import DownloadImage from '../../assets/images/download.svg';
import Dialog from '@material-ui/core/Dialog';
import DialogActions from '@material-ui/core/DialogActions';
import DialogContent from '@material-ui/core/DialogContent';
import DialogContentText from '@material-ui/core/DialogContentText';
import DialogTitle from '@material-ui/core/DialogTitle';
import Button from '@material-ui/core/Button';
import TextField from '@material-ui/core/TextField';
import HelperServices from "../../services"

const columns = [
    { id: 'ReportName', label: 'Test Report Number', minWidth: 170 },
    { id: 'TestEngineerName', label: 'Testing Engineer name', minWidth: 140 },
    {
        id: 'UploadDate',
        label: 'Upload Date',
        minWidth: 170,
        align: 'right',
        format: (value) => value.toLocaleString(),
    },
    {
        id: 'ApprovedSatatus1',
        label: 'Testing Engineer',
        minWidth: 170,
        align: 'right',
        format: (value) => value.toLocaleString(),
    },
    {
        id: 'ApprovedSatatus2',
        label: 'Head of Department',
        minWidth: 170,
        align: 'right',
        format: (value) => value.toFixed(2),
    },
    {
        id: 'ApprovedSatatus3',
        label: 'Business Development Head',
        minWidth: 170,
        align: 'right',
        format: (value) => value.toFixed(2),
    },
    {
        id: 'ReportApprovedSatatus',
        label: 'Status',
        minWidth: 170,
        align: 'right',
        format: (value) => value.toFixed(2),
    },
    {
        id: 'Download',
        label: 'Download',
        minWidth: 70,
        align: 'right',
        format: (value) => value.toFixed(2),
    },
];



const useStyles = makeStyles({
    root: {
        width: '100%',
    },
    container: {
        maxHeight: 440,
    },
});

export default function ApproveReport() {
    const classes = useStyles();
    const [page, setPage] = React.useState(0);
    const [rowsPerPage, setRowsPerPage] = React.useState(10);
    const [open, setOpen] = React.useState(false);
    const [rows, setRows] = React.useState([]);
    const [password, setpassword] = React.useState("");
    const [approvingAuthority, setApprovingAuthority] = React.useState("");
    const [selectedReport, setSelectedReport] = React.useState("");
    const [rejectComment, setrejectComment] = React.useState("");

    React.useEffect(() => {
        HelperServices.getReportFileList()
            .then((data) => {
                setRows(data)
            })
        console.log('mount it!');
    }, []);
    const passwordTyped = (event) => {
        setpassword(event.target.value.toString())
    }
    const rejectCommentTyped = (event) => {
        setrejectComment(event.target.value.toString())
    }
    const handleClickOpen = () => {
        setOpen(true);
    };

    const handleClose = () => {
        setOpen(false);
    };

    const submitPassword = () => {
        if (password == "") {
            alert("Please enter the password")
        }
        else {
            setOpen(false);
            HelperServices.updateStatus(selectedReport, password, approvingAuthority, "Accept", "")
                .then((data) => {
                    if (data == "Status Updated") {
                        HelperServices.getReportFileList()
                            .then((data) => {
                                setRows(data)
                            })
                    }
                    else {
                        alert("Please enter a correct access code")
                    }
                })
                .catch((e) => {

                })
        }
    };
    const rejectReport = () => {
        if (password == "") {
            alert("Please enter the password")
        }
        else {
            if (rejectComment == "") {
                alert("Please enter a comment to reject this report")
            }
            else {
                setOpen(false);
                HelperServices.updateStatus(selectedReport, password, approvingAuthority, "Reject", rejectComment)
                    .then((data) => {
                        if (data == "Status Updated") {
                            HelperServices.getReportFileList()
                                .then((data) => {
                                    setRows(data)
                                })
                        }
                        else {
                            alert("Please enter a correct access code")
                        }
                    })
                    .catch((e) => {

                    })
            }
        }


    };

    const handleChangePage = (event, newPage) => {
        setPage(newPage);
    };

    const handleChangeRowsPerPage = (event) => {
        setRowsPerPage(+event.target.value);
        setPage(0);
    };

    const downloadReport = (ReportRow) => {
        console.log(ReportRow.ReportName)
        HelperServices.downloadFile(ReportRow.ReportName, "FullReport")
    }
    const TableCellClick = (ReportRow, columnId) => {
        if (columnId == "Download") {
            if (ReportRow.ReportApprovedSatatus != "Approved") {
                alert("You can't download a Report till its not approved")
            }
            else {
                downloadReport(ReportRow)
            }
        }
        else {
            if (columnId == "ApprovedSatatus1" || columnId == "ApprovedSatatus2" || columnId == "ApprovedSatatus3") {
                handleClickOpen()
                setSelectedReport(ReportRow.ReportName)
                setApprovingAuthority(columnId)
            }
            else if (columnId == "ReportName") {
                HelperServices.downloadFile(ReportRow.ReportName, "HalfReport")
            }

        }
    }

    return (
        <>
            <Paper className={classes.root}>
                <TableContainer className={classes.container}>
                    <Table stickyHeader aria-label="sticky table">
                        <TableHead>
                            <TableRow>
                                {columns.map((column) => (
                                    <TableCell
                                        key={column.id}
                                        align={column.align}
                                        style={{ minWidth: column.minWidth }}
                                    >
                                        {column.label}
                                    </TableCell>
                                ))}
                            </TableRow>
                        </TableHead>
                        <TableBody>
                            {rows.slice(page * rowsPerPage, page * rowsPerPage + rowsPerPage).map((row) => {
                                return (
                                    <TableRow hover role="checkbox" tabIndex={-1} key={row.code}>
                                        {columns.map((column) => {
                                            const value = row[column.id];
                                            return (
                                                <TableCell key={column.id} align={column.align} onClick={() => TableCellClick(row, column.id)}>
                                                    {
                                                        column.id == "Download" ? <img style={{ height: "25px" }} src={DownloadImage} alt="Download" /> : column.format && typeof value === 'number' ? column.format(value) : value
                                                    }

                                                </TableCell>
                                            );
                                        })}
                                    </TableRow>
                                );
                            })}
                        </TableBody>
                    </Table>
                </TableContainer>
                <TablePagination
                    rowsPerPageOptions={[10, 25, 100]}
                    component="div"
                    count={rows.length}
                    rowsPerPage={rowsPerPage}
                    page={page}
                    onChangePage={handleChangePage}
                    onChangeRowsPerPage={handleChangeRowsPerPage}
                />
            </Paper>
            <Dialog open={open} onClose={handleClose} aria-labelledby="form-dialog-title">
                <DialogTitle id="form-dialog-title">Apporve Report</DialogTitle>
                <DialogContent>
                    <DialogContentText>
                        To Approve this Report, please enter your Secret Code.
          </DialogContentText>
                    <TextField
                        autoFocus
                        margin="dense"
                        id="name"
                        label="Secret Code"
                        type="password"
                        fullWidth
                        onChange={passwordTyped}
                    />
                </DialogContent>
                <DialogContent>
                    <DialogContentText>
                        In case of rejection please enter a comment.
          </DialogContentText>
                    <TextField
                        autoFocus
                        margin="dense"
                        id="Remark"
                        label="Remark"
                        type="text"
                        fullWidth
                        onChange={rejectCommentTyped}

                    />
                </DialogContent>
                <DialogActions>
                    <Button onClick={handleClose} color="primary">
                        Cancel
          </Button>
                    <Button onClick={rejectReport} color="primary">
                        Reject
          </Button>
                    <Button onClick={submitPassword} color="primary">
                        Approve
          </Button>
                </DialogActions>
            </Dialog>
        </>
    );
}
