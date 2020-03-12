import numpy as np
import openpyxl as xl
from copy import deepcopy

class MA_filter:
  def __init__(self, sheet_col_rt, w, s, quantile):
    self.window = w                                 # moving average window size
    self.stride = s                                 # stride
    self.num_atom = 1231663                         # num of platelet particles
    self.perc = quantile                            # quantile: remove outlier data 
    self.masstot = 0.0035768061 * self.num_atom     # platelet total mass

    self.Et = None
    self.Er = None
    self.Omega = None
    self.Ycom = None
    self.rt = None
    self.mag_om = None
    self.raw_t = [x.value for i,x in enumerate(sheet_col_rt) if x.value is not None and i > 0]
    print("number of samples: ", len(self.raw_t))

  def ma_eq(self, nums):
    new_nums = []
    start = end = 0
    current_time = [self.window/2 + self.raw_t[0]]

    for i in range(len(nums)):
      if self.raw_t[i] > current_time[-1] + self.window/2:
        end = i
        while self.raw_t[start] < current_time[-1] - self.window/2:
          start += 1
        
        # compute bounds of 10% and 90% quantile
        h_lim, l_lim = np.percentile(nums[start:end+1], 100-self.perc), np.percentile(nums[start:end+1], self.perc)

        # averaging data within bounds
        new_nums.append(np.average([i for i in nums[start:end] if l_lim<=i<=h_lim]))

        current_time.append(current_time[-1] + self.stride)
    return new_nums, current_time[:-1]

  def read_Et(self, sheet_col_Et):
    if not self.rt:
      self.Et, self.rt = self.ma_eq([x.value*self.num_atom for i,x in enumerate(sheet_col_Et) if x.value is not None and i > 0])
    else:
      self.Et, _ = self.ma_eq([x.value*self.num_atom for i,x in enumerate(sheet_col_Et) if x.value is not None and i > 0])
    self.Et = [i/self.masstot for i in self.Et]

  def read_Er(self, sheet_col_Er):
    if not self.rt:
      self.Er, self.rt = self.ma_eq([x.value*self.num_atom for i,x in enumerate(sheet_col_Er) if x.value is not None and i > 0])
    else:
      self.Er, _ = self.ma_eq([x.value*self.num_atom for i,x in enumerate(sheet_col_Er) if x.value is not None and i > 0])

  def read_Omega(self, sheet_col_Omegax, sheet_col_Omegay, sheet_col_Omegaz):
    if not self.rt:
      self.Omegax, self.rt = self.ma_eq([x.value for i,x in enumerate(sheet_col_Omegax) if x.value is not None and i > 0])
      self.Omegay, _ = self.ma_eq([x.value for i,x in enumerate(sheet_col_Omegay) if x.value is not None and i > 0])
      self.Omegaz, _ = self.ma_eq([x.value for i,x in enumerate(sheet_col_Omegaz) if x.value is not None and i > 0])
    else:
      self.Omegax, _ = self.ma_eq([x.value for i,x in enumerate(sheet_col_Omegax) if x.value is not None and i > 0])
      self.Omegay, _ = self.ma_eq([x.value for i,x in enumerate(sheet_col_Omegay) if x.value is not None and i > 0])
      self.Omegaz, _ = self.ma_eq([x.value for i,x in enumerate(sheet_col_Omegaz) if x.value is not None and i > 0])
    self.mag_om = np.sqrt([i*i+j*j+k*k for i,j,k in zip([i*1e3/2.0833 for i in self.Omegax], [i*1e3/2.0833 for i in self.Omegay], [i*1e3/2.0833 for i in self.Omegaz])])

  def read_Ycom(self, sheet_col_Ycom):
    if not self.rt:
      self.Ycom, self.rt = self.ma_eq([x.value for i,x in enumerate(sheet_col_Ycom) if x.value is not None and i > 0])
    else:
      self.Ycom, _ = self.ma_eq([x.value for i,x in enumerate(sheet_col_Ycom) if x.value is not None and i > 0])


wb1 = xl.load_workbook("drive/My Drive/ATS/ats-nn/training.xlsx", data_only=True)
print(wb1.sheetnames)
f2_50 = wb1['f2-50']

MA = MA_filter(f2_50['ac'], 0.02, 0.005, 10)
MA.read_Et(f2_50['j'])
MA.read_Er(f2_50['k'])
MA.read_Omega(f2_50['s'], f2_50['t'], f2_50['u'])
MA.read_Ycom(f2_50['w'])