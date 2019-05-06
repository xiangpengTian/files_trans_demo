package com.txp.files.controller;

import com.txp.files.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
public class ExcelUpAndDown {
    @RequestMapping("upExcel")
    @ResponseBody
    public Map<String, Object> readExcel(@RequestParam(value = "file") MultipartFile excelFile, HttpServletRequest req, HttpServletResponse resp) {
        Map<String, Object> param = new HashMap<String, Object>();
//        List<Car4SInfo> infoList = new ArrayList<Car4SInfo>();
        try {
            List<String[]> rowList = ExcelUtils.readExcel(excelFile);
            for (int i = 0; i < rowList.size(); i++) {
                String[] row = rowList.get(i);
              /*  Car4SInfo info = new Car4SInfo();
                info.setDealers(row[0]);
                info.setProvince(row[1]);
                info.setCity(row[2]);
                info.setAddress(row[3]);
                info.setPhone(row[4]);
                info.setCar_model(row[5]);
                service.insertCar4SInfo(info);*/
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        param.put("dataStatus", 1);
        return param;
    }

    /**
     * 下载模板
     *
     * @param request
     * @param response
     */
    @RequestMapping("downloadTmpl.do")
    public void downloadTmpl(HttpServletRequest request, HttpServletResponse response) {
        try {
            String fileName = "car4sinfo.xlsx";
            String path = request.getSession().getServletContext().getRealPath("/template/4sinfo") + "/" + fileName;
            path = path.replace("\\", "/");
            File file = new File(path);
            String filename = file.getName();
            // 取得文件的后缀名。
            String ext = filename.substring(filename.lastIndexOf(".") + 1).toUpperCase();
            // 以流的形式下载文件。
            InputStream fis = new BufferedInputStream(new FileInputStream(path));
            byte[] buffer = new byte[fis.available()];
            fis.read(buffer);
            fis.close();
            // 清空response
            response.reset();
            // 设置response的Header
            response.addHeader("Content-Disposition", "attachment;filename=" + new String(filename.getBytes()));
            response.addHeader("Content-Length", "" + file.length());
            OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
            response.setContentType("application/octet-stream");
            toClient.write(buffer);
            toClient.flush();
            toClient.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

}