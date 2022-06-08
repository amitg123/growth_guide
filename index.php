<?php
session_start();
$con=mysqli_connect('localhost','root','','import_excle');
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (isset($_POST['submit'])) {
    $allowed_ext = ['xls', 'csv', 'xlsx'];
    $fileName = $_FILES['doc']['name'];
    $checking = explode(".", $fileName);
    $file_ext = end($checking);
    if (in_array($file_ext, $allowed_ext)) {
        $targetPath = $_FILES['doc']['tmp_name'];
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($targetPath);
        $data =$spreadsheet->getActiveSheet()->toArray();
        foreach($data as $row){
             $id = $row['0'];
             $f_name=$row['1'];
             $l_name=$row['2'];
             $mobile=$row['3'];
             $checkStudent ="SELECT id FROM user WHERE id='$id'";
             $checkStudent_result = mysqli_query($con,$checkStudent);
             if(mysqli_num_rows($checkStudent_result)>0){
                 $up_query ="UPDATE user SET first_name='$f_name',last_name='$l_name',mobile='$mobile' WHERE id='$id'";
                 $up_result= mysqli_query($con, $up_query);
                 $msg=1;
             }else{
                 $in_query ="INSERT INTO user (first_name,last_name,mobile) VALUES ('$f_name','$l_name','$mobile')";
                 $in_result=mysqli_query($con,$in_query);
                 $msg=1;
             }
        }
        if($msg){
            $_SESSION['status'] = "file Imported Successfully";
        header("Location: index.php");
        }else{
            $_SESSION['status'] = "File Imported Fail";
        header("Location: index.php");
        }
    } else {
        $_SESSION['status'] = "Invalid File";
        header("Location: index.php");
        exit(0);
    }
}
?>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Growth Study</title>

    <link rel="stylesheet" href="assets/css/style.css">
    <link rel="stylesheet" type="text/css" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.1.1/css/all.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.8.3/font/bootstrap-icons.css">
    <!-- CSS only -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.0-beta1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-0evHe/X+R7YkIZDRvuzKMRqM+OrBnVFBL6DOitfPri4tjfHxaWutUpFmBp4vmVor" crossorigin="anonymous">
    <!-- JavaScript Bundle with Popper -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.2.0-beta1/dist/js/bootstrap.bundle.min.js" integrity="sha384-pprn3073KE6tl6bjs2QrFaJGz5/SUsLqktiwsUTF55Jfv3qYSDhgCecCxMW52nD2" crossorigin="anonymous"></script>
</head>

<body>
    <header>
        <nav class="navbar bg-light">
            <div class="container">
                <a class="navbar-brand" href="#"><img src="assets/img/logo-1-300x180.png" width="100px" height="50px"></a>
                <button class="navbar-toggler" type="button" data-bs-toggle="offcanvas" data-bs-target="#offcanvasNavbar" aria-controls="offcanvasNavbar">
                    <span class="navbar-toggler-icon"></span>
                </button>
                <div class="offcanvas offcanvas-end" tabindex="-1" id="offcanvasNavbar" aria-labelledby="offcanvasNavbarLabel">
                    <div class="offcanvas-header" style="border-bottom:solid 1px red ;">
                        <h5 class="offcanvas-title" id="offcanvasNavbarLabel"><img src="assets/img/logo-1-300x180.png" width="100px" height="50px"></h5>
                        <button type="button" class="btn-close" data-bs-dismiss="offcanvas" aria-label="Close"></button>
                    </div>
                    <div class="offcanvas-body">
                        <ul class="navbar-nav justify-content-end flex-grow-1 pe-3">
                            <li class="nav-item" style="font-weight:bold ;">
                                <a class="nav-link active" aria-current="page" href="#"><i class="bi bi-house-heart-fill"></i> &nbsp;Home</a>
                            </li>
                            <hr>
                            <li class="nav-item" style="font-weight:bold ;">
                                <a class="nav-link active" aria-current="page" href="#"><i class="bi bi-blockquote-left"></i>&nbsp;Blog</a>
                            </li>
                            <hr>
                            <li class="nav-item dropdown" style="font-weight:bold ;">
                                <a class="nav-link active dropdown-toggle" href="#" id="offcanvasNavbarDropdown" role="button" data-bs-toggle="dropdown" aria-expanded="false">
                                    <i class="bi bi-person-lines-fill"></i>&nbsp;Contact Us
                                </a>
                                <ul class="dropdown-menu" aria-labelledby="offcanvasNavbarDropdown">
                                    <li><a class="dropdown-item" href="#"><i class="bi bi-telephone-fill"></i>&nbsp;+91-9267XXXXXX</a></li>
                                    <li><a class="dropdown-item" href="#"><i class="bi bi-telephone-fill"></i>&nbsp;+91-9267XXXXXX</a></li>
                                    <li>
                                        <hr class="dropdown-divider">
                                    </li>
                                    <li><a class="dropdown-item" href="#"><i class="bi bi-envelope-open-fill"></i>&nbsp;amitg2096@gmail.com</a></li>
                                </ul>
                            </li>
                            <hr>
                            <li class="nav-item" style="font-weight:bold ;">
                                <a class="nav-link active" href="#"><i class="bi bi-life-preserver"></i>&nbsp;Services</a>
                            </li>
                            <hr>
                            <li class="nav-item dropdown" style="font-weight:bold ;">
                                <a class="nav-link active dropdown-toggle" href="#" id="offcanvasNavbarDropdown" role="button" data-bs-toggle="dropdown" aria-expanded="false">
                                    <i class="bi bi-person-circle"></i> &nbsp;Account
                                </a>
                                <ul class="dropdown-menu" aria-labelledby="offcanvasNavbarDropdown">
                                    <li><a class="dropdown-item" href="#"><i class="bi bi-person-check"></i>&nbsp;Account Info</a></li>
                                    <li><a class="dropdown-item" href="#"><i class="bi bi-gear"></i>&nbsp;Account Setting</a></li>
                                    <li>
                                        <hr class="dropdown-divider">
                                    </li>
                                    <li><a class="dropdown-item" href="#"><i class="bi bi-box-arrow-right"></i>&nbsp;Logout</a></li>
                                </ul>
                            </li>
                            <hr>
                        </ul>
                    </div>
                </div>
            </div>
        </nav>
    </header>
    <!-- form -->
    <div class="main-content clear">
        <div class="section__content section__content--p30">
            <div class="container-fluid">
                <div class="row justify-content-md-center mt-5 mb-5">
                    <div class="col-lg-6">
                        <div class="card">
                            <div class="card-header">
                                <strong>Import Excel File</strong>&nbsp;&nbsp;(<a href="assets/sample/StudentRegisteration.xlsx">Download Sample</a>)
                            </div>
                            <div class="card-body card-block">
                                <?php
                                if(isset($_SESSION['status'])){
                                    ?>
                                    <label class="mb-2" style="color:red;"><?php echo $_SESSION['status']; unset($_SESSION['status']);?></label> 
                                    <?php
                                }
                                ?>
                                <form method="post" enctype="multipart/form-data">
                                    <div class="has-success form-group">
                                        <label for="inputIsValid" class=" form-control-label mb-2">Select Your Excel File</label>
                                        <input type="file" name="doc" class="form-control">
                                    </div>
                                    <div class="form-group text-center mt-4">
                                        <input type="submit" name="submit" value="Submit" class=" btn btn-success">
                                    </div>
                                </form>
                            </div>
                        </div>
                    </div>
                </div>

            </div>
        </div>
    </div>
    <footer class="footer">
        <div class="container">
            <div class="row justify-content-md-center">
                <div class="footer-col col-lg-3">
                    <h4>Company</h4>
                    <ul class="list-group">
                        <li><a href="#">About Us</a></li>
                        <li><a href="#">Our Services</a></li>
                        <li><a href="#">Privacy Policy</a></li>
                        <li><a href="#">Affiliate Program</a></li>
                    </ul>
                </div>
                <div class="footer-col col-lg-3">
                    <h4>Get Help</h4>
                    <ul class="list-group">
                        <li><a href="#">FAQ</a></li>
                        <li><a href="#">Returns</a></li>
                        <li><a href="#">Order Status</a></li>
                        <li><a href="#">Payment Option</a></li>
                    </ul>
                </div>
                <div class="footer-col col-lg-3">
                    <h4>Services</h4>
                    <ul class="list-group">
                        <li><a href="#">Content Marketing</a></li>
                        <li><a href="#">Digital Marketing</a></li>
                        <li><a href="#">Creative Writing</a></li>
                        <li><a href="#">Website Development</a></li>
                    </ul>
                </div>
                <div class="footer-col col-lg-3">
                    <h4>Follow Us</h4>
                    <div class="social-links">
                        <a href="#"><i class="fab fa-facebook-f"></i></a>
                        <a href="#"><i class="fab fa-twitter"></i></a>
                        <a href="#"><i class="fab fa-instagram"></i></a>
                        <a href="#"><i class="fab fa-linkedin-in"></i></a>
                    </div>
                </div>
            </div>
        </div>
    </footer>
</body>

</html>