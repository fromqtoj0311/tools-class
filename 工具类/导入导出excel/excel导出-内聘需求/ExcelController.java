package com.jd.position.controller;

import com.jd.jsf.gd.util.JsonUtils;
import com.jd.performance.util.LoginUtils;
import com.jd.position.dao.ChooseUserDao;
import com.jd.position.dao.LevelDao;
import com.jd.position.excel.ExcelData;
import com.jd.position.excel.ExportExcelUtils;
import com.jd.position.message.ChooseResult;
import com.jd.position.message.ChooseState;
import com.jd.position.model.ChooseUserVo;
import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
/**
 * 功能描述:
 *
 * @return: 导出模版
 * @auther:
 * @date: 2018/6/19 16:27
 */
@Controller
@RequestMapping("/excel")
public class ExcelController {
  private static final Logger logger = Logger.getLogger(ExcelController.class);

  @Autowired private ChooseUserDao chooseUserDao;
  @Autowired private LevelDao newLevelDao;

  @Value("${safetyKey_Choose}")
  private String safetyKey;

  @RequestMapping("/export")
  @ResponseBody
  public String excel(HttpServletRequest request, HttpServletResponse response) throws Exception {

    String emplErp = null;
    try {
      emplErp = LoginUtils.getERPForJDME(request, safetyKey);
    } catch (Exception e) {
      e.printStackTrace();
    }
    if (!"qijian27".equals(emplErp)) {
      return successMsg();
    }

    ExcelData data = new ExcelData();
    data.setName("员工竞聘表");

    List<String> titles = gettitles();
    data.setTitles(titles);

    List<ChooseUserVo> chooseUser = chooseUserDao.selectAllChooseUser();
    System.out.println(chooseUser);
    if (chooseUser == null || chooseUser.size() == 0) {
      return FailMsg();
    }
    CreateExcel(response, data, chooseUser);
    return successMsg();
  }

  @RequestMapping("/exportByTime")
  @ResponseBody
  public String excel(
      HttpServletRequest request,
      HttpServletResponse response,
      @RequestParam("chooseTime") String chooseTime)
      throws Exception {
    String emplErp = null;
    try {
      emplErp = LoginUtils.getERPForJDME(request, safetyKey);
    } catch (Exception e) {
      e.printStackTrace();
    }
    if (!"qijian27".equals(emplErp)) {
      return successMsg();
    }

    ExcelData data = new ExcelData();
    data.setName("员工竞聘表");

    List<String> titles = gettitles();

    List<ChooseUserVo> chooseUser = chooseUserDao.selectChooseTimeChooseUser(chooseTime);
    System.out.println(chooseUser);
    if (chooseUser == null || chooseUser.size() == 0) {
      return FailMsg();
    }

    CreateExcel(response, data, chooseUser);
    return successMsg();
  }

  private List<String> gettitles() {
    return Arrays.asList(
        "ID", "ERP编码", "姓名", "岗位名称", "职级", "部门全路径名称", "应聘部门", "应聘岗位", "期望职级", "员工邮箱", "时间戳");
  }

  private void CreateExcel(HttpServletResponse response, ExcelData data, List<ChooseUserVo> list0)
      throws Exception {
    List<List<Object>> rowss = new ArrayList();
    for (int i = 0; i <= list0.size() - 1; i++) {
      ChooseUserVo vo = list0.get(i);
      List<Object> row = new ArrayList();
      row.add(i + 1);
      row.add(vo.getEmplErp());
      row.add(vo.getEmplName());
      // row.add(vo.getPositionCode());
      row.add(vo.getPositionName());
      row.add(vo.getLevelName());
      row.add(vo.getOrganizationFullName());
      row.add(vo.getSelectOrganization());
      row.add(vo.getSelectPosition());
      row.add(vo.getSelectLevelName());
      row.add(vo.getEmail());
      row.add(vo.getTime().toString().substring(0, 19));
      rowss.add(row);
    }
    data.setRows(rowss);

    ExportExcelUtils.exportExcel(response, "内部竞聘.xlsx", data);
  }

  @RequestMapping("/delMsg")
  @ResponseBody
  public void delMsg(HttpServletRequest request, @RequestParam String erp) {
    String emplErp = null;
    try {
      emplErp = LoginUtils.getERPForJDME(request, safetyKey);
    } catch (Exception e) {
      e.printStackTrace();
    }
    if ("qijian27".equals(emplErp)) {
      newLevelDao.delMsg(erp);
    }
  }

  private String successMsg() {
    return JsonUtils.toJSONString(ChooseResult.fait(ChooseState.EXCELSUCCESS));
  }

  private String FailMsg() {
    return JsonUtils.toJSONString(ChooseResult.fait(ChooseState.EXCELFAIL));
  }
}
