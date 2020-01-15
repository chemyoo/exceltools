package com.pers.chemyoo.smb;
 
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.EnumSet;
import java.util.List;
import java.util.concurrent.TimeUnit;

import com.hierynomus.msdtyp.AccessMask;
import com.hierynomus.msfscc.fileinformation.FileIdBothDirectoryInformation;
import com.hierynomus.mssmb2.SMB2CreateDisposition;
import com.hierynomus.mssmb2.SMB2ShareAccess;
import com.hierynomus.smbj.SMBClient;
import com.hierynomus.smbj.SmbConfig;
import com.hierynomus.smbj.auth.AuthenticationContext;
import com.hierynomus.smbj.connection.Connection;
import com.hierynomus.smbj.session.Session;
import com.hierynomus.smbj.share.DiskShare;
import com.hierynomus.smbj.share.File;
 
/**
 * SMB2连接示例
 * 例: 我们当前要连接 IP为:123.123.123.123 目录为: SRC/SMB2/ 下的文件 
 * @author liuyb
 *
 */
public class SMB2Connect {
	private static final String SHARE_DOMAIN = "";
	private static final String SHARE_USER = "Administrator";
	private static final String SHARE_PASSWORD = "chemyoo";
	private static final String SHARE_SRC_DIR = "bwdata";
 
	public static void main(String[] args) {
		// 设置超时时间(可选)
		SmbConfig config = SmbConfig.builder().withTimeout(120, TimeUnit.SECONDS)
				.withTimeout(120, TimeUnit.SECONDS) // 超时设置读，写和Transact超时（默认为60秒）
	            .withSoTimeout(180, TimeUnit.SECONDS) // Socket超时（默认为0秒）
	            .build();
		
		// 如果不设置超时时间	SMBClient client = new SMBClient();
		SMBClient client = new SMBClient(config);
 
		try (Connection connection = client.connect("localhost")){
			AuthenticationContext ac = new AuthenticationContext(SHARE_USER, SHARE_PASSWORD.toCharArray(), SHARE_DOMAIN);
			Session session = connection.authenticate(ac);
 
			// 连接共享文件夹
			DiskShare share = (DiskShare) session.connectShare(SHARE_SRC_DIR);
			
			String folder = SHARE_SRC_DIR;
			String dstRoot = "D:/download/";	// 如: D:/smd2/
			List<FileIdBothDirectoryInformation> list = share.list("upload/ZJWJ", "*.zip");
			for (FileIdBothDirectoryInformation f : list) {
				String fileName = f.getFileName();
				if(fileName.equals(".") || fileName.equals("..")) {
					continue;
				}
				String filePath = folder + "/" + fileName;
				String dstPath = dstRoot + "/" + fileName;
				java.io.File destFile = new java.io.File(dstPath);
				if(!destFile.getParentFile().exists()) {
					destFile.getParentFile().mkdirs();
				}
				if(destFile.exists()) {
					destFile.delete();
					destFile.createNewFile();
				}
					
				FileOutputStream fos = new FileOutputStream(dstPath);
				BufferedOutputStream bos = new BufferedOutputStream(fos);
				
				if (share.fileExists(filePath)) {
					System.out.println("正在下载文件:" + f.getFileName());
					
					File smbFileRead = share.openFile(filePath, EnumSet.of(AccessMask.GENERIC_READ), null, SMB2ShareAccess.ALL, SMB2CreateDisposition.FILE_OPEN, null);
					InputStream in = smbFileRead.getInputStream();
					byte[] buffer = new byte[4096];
					int len = 0;
					while ((len = in.read(buffer, 0, buffer.length)) != -1) {
						bos.write(buffer, 0, len);
					}
					
					bos.flush();
					bos.close();
					
					System.out.println("文件下载成功");
					System.out.println("==========================");
				} else {
					System.out.println(dstPath + "文件不存在");
				}
            }
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (client != null) {
				client.close();
			}
		}
	}
}
