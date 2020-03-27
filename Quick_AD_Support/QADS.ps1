Param
 (
	[String]$Restart	
 )

 
 
# Get-ADDefaultDomainPasswordPolicy -Current LoggedOnUser 
 
 
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 		 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MahApps.Metro.dll')       				| out-null 

[System.Windows.Forms.Application]::EnableVisualStyles()
 
 
If ($Restart -ne "") 
	{
		sleep 10
	}

$Global:Current_Folder = $PSScriptRoot
If (test-path "C:\Windows\System32\ServerManager.exe")
	{
		$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\ServerManager.exe")	
	}
ElseIf (test-path "C:\Windows\System32\ServerManager.exe")
	{
		$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\mmc.exe")	
	}
	
$Current_Version = "1.1"



########################################################################################################################################################################################################	
#*******************************************************************************************************************************************************************************************************
#																						DIFFERENT GUI DISPLAY PART
#*******************************************************************************************************************************************************************************************************
########################################################################################################################################################################################################


# ----------------------------------------------------
# Part - User GUI
# ----------------------------------------------------

[xml]$XAML_Users =  
'
<Controls:MetroWindow 
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	xmlns:i="http://schemas.microsoft.com/expression/2010/interactivity"		
	xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
	Title="Quick AD Support - User Analysis - Part v1.1" 
	Width="470" 
	ResizeMode="NoResize"
	Height="Auto" 
	SizeToContent="Height" 			
	BorderBrush="DodgerBlue"
	BorderThickness="0.5"
	WindowStartupLocation ="CenterScreen"
	>

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="C:\ProgramData\Quick_AD_Support\QADS_Systray\resources\Icons.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Cobalt.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>

   <Controls:MetroWindow.LeftWindowCommands>
        <Controls:WindowCommands>
            <Button>
                <StackPanel Orientation="Horizontal">
                    <Rectangle Width="15" Height="15" Fill="{Binding RelativeSource={RelativeSource AncestorType=Button}, Path=Foreground}">
                        <Rectangle.OpacityMask>
                            <VisualBrush Stretch="Fill" Visual="{StaticResource appbar_user}" />							
                        </Rectangle.OpacityMask>
                    </Rectangle>					
                </StackPanel>
            </Button>				
        </Controls:WindowCommands>	
    </Controls:MetroWindow.LeftWindowCommands>		

    <Grid>	
        <StackPanel Orientation="Vertical" Margin="0,0,0,0">
	
		<!--	<StackPanel VerticalAlignment="Center">
				<Image Width="150" Height="70" Source="C:\ProgramData\Quick_AD_Support\QADS_Systray\logo.png" ></Image>	
			</StackPanel>	-->
			
			<StackPanel VerticalAlignment="Center">
				<Image Width="150" Height="70" Source="C:\ProgramData\Quick_AD_Support\QADS_Systray\logo.jpg" ></Image>	
			</StackPanel>				

			<StackPanel Margin="0,5,0,0" >				
				<StackPanel Name="Expander_Selection_Block">							
					<Expander Name="Expander_Selection" Header="User selection"  Margin="0,0,0,0"  IsExpanded="False" Height="auto">  
						<Grid Background="white" >
							<StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Orientation="Horizontal">	
								<Label Margin="0,0,0,0" Content="Type user name" Foreground="Black" FontSize="18"/>
								<StackPanel  Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">												
									<TextBox ToolTipService.ToolTip="Type the user name you are looking for" Name="User_TxtBox" Width="120" FontSize="16"></TextBox>
									<Button Width="40" Name="Check_User_btn" BorderThickness="0" Margin="0,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#2196f3">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_magnify}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>	
									
									<StackPanel Name="User_Block_OK">
										<Button Width="40" Name="user_OK" BorderThickness="0" Margin="5,0,0,0" 
											Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#00a300">
											<Rectangle Width="20" Height="20"  Fill="white" >
												<Rectangle.OpacityMask>
													<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_check}"/>
												</Rectangle.OpacityMask>
											</Rectangle>
										</Button>	
									</StackPanel>
									
									<StackPanel Name="User_Block_KO">
										<Button Width="40" x:Name="User_KO" BorderThickness="0" Margin="5,0,0,0" 
											Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="Red">
											<Rectangle Width="20" Height="20"  Fill="white" >
												<Rectangle.OpacityMask>
													<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_warning}"/>
												</Rectangle.OpacityMask>
											</Rectangle>
										</Button>	
									</StackPanel>
								</StackPanel>				
							</StackPanel>									
						</Grid>
					</Expander> 
				</StackPanel>									
				
				<StackPanel Name="Global_Warning_Block">
					<Expander Name="Expander_Warning" Header="Warning infos"  Margin="0,0,0,0" IsExpanded="False" Height="auto">  
						<Grid Background="white">
							<StackPanel Orientation="Vertical">				
						
								<StackPanel Orientation="Vertical" Margin="0,0,0,0" HorizontalAlignment="Center" VerticalAlignment="Center">				
									<Label Name="Warning_Label" />
								</StackPanel>					

								<StackPanel Name="DatagGrid_MultiUsers" Margin="0,10,0,0" HorizontalAlignment="Center" VerticalAlignment="Center">
									<Grid Background="CornFlowerBlue">								
										<DataGrid IsReadOnly="True" RowHeaderWidth="0" SelectionMode="Single" Height="100" AutoGenerateColumns="False" Name="DatagGrid_Users" ItemsSource="{Binding}" Margin="1,1,1,1" >										
											<DataGrid.Columns>										
												<DataGridTextColumn Width="auto" Header="Computer Account" Binding="{Binding Account}"/>							
												<DataGridTextColumn Width="auto" Header="Display name" Binding="{Binding Name}"/>							
											</DataGrid.Columns>
										</DataGrid>	
									</Grid>
								</StackPanel>		
							</StackPanel>													
						</Grid>
					</Expander> 				
				</StackPanel>
				
			
				<StackPanel Name="Expander_Basic_Block">				
					<Expander Name="Expander_Basic" Header="Basic infos"  Margin="0,0,0,0" Background="gray" IsExpanded="False" Height="auto"> 
						<ScrollViewer  CanContentScroll="True" Height="100">        					
							<Grid Background="white" HorizontalAlignment="Left">
								<StackPanel Orientation="Vertical" Margin="0,0,0,0">	
									<StackPanel Orientation="Horizontal">
										<Label Content="Is this account enabled ?"/>
										<Label Name="IsEnabled_Value"/>				
									</StackPanel>								
								
									<StackPanel Orientation="Horizontal">
										<Label Content="Is this account locked ?"/>
										<Label Name="Locked_Value"/>				
									</StackPanel>	
									
									<StackPanel Orientation="Horizontal">
										<Label Content="Is password expired ?"/>
										<Label Name="IsPWDExpired_Value"/>				
									</StackPanel>									
									
									<StackPanel Orientation="Horizontal">
										<Label Content="Display name:"/>
										<Label Name="DisplayName_Value"/>				
									</StackPanel>	
								
									<StackPanel Orientation="Horizontal">
										<Label Content="Logon account:"/>
										<Label Name="Account_Value"/>				
									</StackPanel>	
									
									<StackPanel Orientation="Horizontal">
										<Label Content="Password last change:"/>
										<Label Name="PWDLastChange_Value"/>				
									</StackPanel>									
									
									<StackPanel Orientation="Horizontal">
										<Label Content="Password expiration date:"/>
										<Label Name="PWD_Expiration_Date_Value"/>				
									</StackPanel>			

									<StackPanel Orientation="Horizontal">
										<Label Content="Last logon date:"/>
										<Label Name="LastLogOn_Value"/>				
									</StackPanel>									

									<StackPanel Orientation="Horizontal">
										<Label Content="User OU:"/>
										<TextBox Name="UserOU_Value" Width="400"/>				
									</StackPanel>							
								</StackPanel>									
							</Grid>
						</ScrollViewer> 										
					</Expander> 
				</StackPanel>


				<StackPanel Name="Expander_More_Block">								
					<Expander Name="Expander_More" Header="More informations" Margin="0,0,0,0" IsExpanded="False" Height="auto">  
						<ScrollViewer HorizontalScrollBarVisibility="Auto" CanContentScroll="True" Height="100">        
							<Grid Background="white" HorizontalAlignment="Left" >
								<StackPanel Orientation="Vertical" Margin="0,0,0,0">						
									<StackPanel Orientation="Horizontal">
										<Label Content="User department:"/>
										<Label Name="Dept_Value"/>				
									</StackPanel>

									<StackPanel Orientation="Horizontal">
										<Label Content="Mail:"/>
										<Label Name="Mail_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Office:"/>
										<Label Name="Office_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Last bad password attempt:"/>
										<Label Name="LastBadPWD_Value"/>				
									</StackPanel>									

									<StackPanel Orientation="Horizontal">
										<Label Content="Account created on:"/>
										<Label Name="WhenCreated_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Account changed on:"/>
										<Label Name="WhenChanged_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Can user change password ?"/>
										<Label Name="CannotChangePassword_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Password never expires ?"/>
										<Label Name="PWDNeverExpires_Value"/>				
									</StackPanel>									
								</StackPanel>	
							</Grid>							
						</ScrollViewer> 											
					</Expander> 
				</StackPanel>	

				<StackPanel Name="Expander_More_options_Block">												
					<Expander Name="Expander_More_options" Header="Other options"  Margin="0,0,0,0" Background="gray" IsExpanded="False" Height="auto">  
						<Grid Background="white" HorizontalAlignment="Center">
							<StackPanel Orientation="Vertical" Margin="0,0,0,0">						
								<StackPanel Orientation="Horizontal">
								
									<Button Width="40" ToolTip="Export User properties to a CSV file" Name="Export_User_Values" BorderThickness="0" Margin="0,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#2196f3">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_people_profile}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>	
								
									<Button Width="40" ToolTip="Unlock this account" Name="User_Unlock_Account" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#00a300">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_unlock}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>	
									
									<Button Width="40" ToolTip="Enable this account" Name="User_Enable_Account" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="black">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_power}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>

									<Button Width="40" ToolTip="Delete this account" Name="Delete_User_Account" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="red">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_delete}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>									

									<Button Width="40" ToolTip="Change the user password" Name="Change_Password" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#603cba">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_key}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>		

									<Button Width="40" ToolTip="Diplay in which group the user belongs" Name="User_Display_Groups" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#2b5797">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_group}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>	
								</StackPanel>
							</StackPanel>						
						</Grid>
					</Expander> 				
				</StackPanel>	
				
				<StackPanel Name="Expander_GroupMembership_Block">
					<Expander Name="Expander_GroupMembership" Header="Group membership"  Margin="0,0,0,0" IsExpanded="False" Height="auto">  
						<Grid Background="white">
							<StackPanel Orientation="Vertical">							
								<DataGrid HeadersVisibility="Row" IsReadOnly="True" RowHeaderWidth="0" SelectionMode="Single" Height="100" AutoGenerateColumns="False" Name="DatagGrid_Groups" ItemsSource="{Binding}" Margin="2,2,2,2" >										
									<DataGrid.Columns>										
										<DataGridTextColumn Width="auto" Header="Name" Binding="{Binding Name}"/>							
									</DataGrid.Columns>
								</DataGrid>	
							</StackPanel>													
						</Grid>
					</Expander> 				
				</StackPanel>					
				
				<StackPanel Name="Expander_Change_Password_Block">												
					<Expander Name="Expander_Change_Password" Header="Change user password"  Margin="0,0,0,0"  IsExpanded="False" Height="auto">  
						<Grid Background="white" HorizontalAlignment="Center">
							<StackPanel Orientation="Vertical" Margin="0,0,0,0">	
								<StackPanel  Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">												
									<PasswordBox  
									Name="First_Password"  Width="120"  
									Controls:TextBoxHelper.ClearTextButton="{Binding RelativeSource={RelativeSource Self}, Path=(Controls:TextBoxHelper.HasText), Mode=OneWay}" 
									Controls:TextBoxHelper.IsWaitingForData="True" 
									Controls:TextBoxHelper.Watermark="New password" 	
									Style="{StaticResource MetroButtonRevealedPasswordBox}"								
									/>											
									
									<PasswordBox  
									Name="Confirm_Password"  Width="120" Margin="5,0,0,0" 
									Controls:TextBoxHelper.ClearTextButton="{Binding RelativeSource={RelativeSource Self}, Path=(Controls:TextBoxHelper.HasText), Mode=OneWay}" 
									Controls:TextBoxHelper.IsWaitingForData="True" 
									Controls:TextBoxHelper.Watermark="Confirm password" 	
									Style="{StaticResource MetroButtonRevealedPasswordBox}"								
									/>												
									
									<Button Width="40" Name="Set_Password" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#2196f3">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_key_old}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>	
								
								</StackPanel>	
							</StackPanel>						
						</Grid>
					</Expander> 				
				</StackPanel>					
			</StackPanel>	
        </StackPanel>		
    </Grid>
</Controls:MetroWindow>        
'
$Users_Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAML_Users))


# Selection Part
$Expander_Selection_Block = $Users_Window.findname("Expander_Selection_Block") 
$Expander_Selection = $Users_Window.findname("Expander_Selection") 
$User_TxtBox = $Users_Window.findname("User_TxtBox") 
$Check_User_btn = $Users_Window.findname("Check_User_btn") 
$User_Block_OK = $Users_Window.findname("User_Block_OK") 
$User_Block_KO = $Users_Window.findname("User_Block_KO") 
$user_OK = $Users_Window.findname("Expander_More") 
$user_KO = $Users_Window.findname("Expander_More") 

# Warning
$Global_Warning_Block = $Users_Window.findname("Global_Warning_Block") 
$Expander_Warning = $Users_Window.findname("Expander_Warning") 
$Warning_Label = $Users_Window.findname("Warning_Label") 
$DatagGrid_MultiUsers = $Users_Window.findname("DatagGrid_MultiUsers") 
$DatagGrid_Users = $Users_Window.findname("DatagGrid_Users") 


# Basic infos block
$Expander_Basic_Block = $Users_Window.findname("Expander_Basic_Block") 
$Expander_Basic = $Users_Window.findname("Expander_Basic") 
$IsEnabled_Value = $Users_Window.findname("IsEnabled_Value") 
$Locked_Value = $Users_Window.findname("Locked_Value") 
$IsPWDExpired_Value = $Users_Window.findname("IsPWDExpired_Value") 
$DisplayName_Value = $Users_Window.findname("DisplayName_Value") 
$Account_Value = $Users_Window.findname("Account_Value") 
$PWDLastChange_Value = $Users_Window.findname("PWDLastChange_Value") 
$PWD_Expiration_Date_Value = $Users_Window.findname("PWD_Expiration_Date_Value") 
$LastLogOn_Value = $Users_Window.findname("LastLogOn_Value") 
$UserOU_Value = $Users_Window.findname("UserOU_Value") 


# More infos block
$Expander_More_Block = $Users_Window.findname("Expander_More_Block") 
$Expander_More = $Users_Window.findname("Expander_More") 
$Dept_Value = $Users_Window.findname("Dept_Value") 
$Mail_Value = $Users_Window.findname("Mail_Value") 
$Office_Value = $Users_Window.findname("Office_Value") 
$LastBadPWD_Value = $Users_Window.findname("LastBadPWD_Value") 
$WhenCreated_Value = $Users_Window.findname("WhenCreated_Value") 
$WhenChanged_Value = $Users_Window.findname("WhenChanged_Value") 
$CannotChangePassword_Value = $Users_Window.findname("CannotChangePassword_Value") 
$PWDNeverExpires_Value = $Users_Window.findname("PWDNeverExpires_Value") 


# More options action block
$Expander_More_options_Block = $Users_Window.findname("Expander_More_options_Block") 
$Expander_More_options = $Users_Window.findname("Expander_More_options") 
$Export_User_Values = $Users_Window.findname("Export_User_Values") 
$User_Unlock_Account = $Users_Window.findname("User_Unlock_Account") 
$Change_Password = $Users_Window.findname("Change_Password") 
$User_Display_Groups = $Users_Window.findname("User_Display_Groups") 
$Delete_User_Account = $Users_Window.findname("Delete_User_Account") 


# Group Membership block
$Expander_GroupMembership_Block = $Users_Window.findname("Expander_GroupMembership_Block") 
$Expander_GroupMembership = $Users_Window.findname("Expander_GroupMembership") 
$DatagGrid_Groups = $Users_Window.findname("DatagGrid_Groups") 


# Change password block
$Expander_Change_Password_Block = $Users_Window.findname("Expander_Change_Password_Block") 
$Expander_Change_Password = $Users_Window.findname("Expander_Change_Password") 
$First_Password = $Users_Window.findname("First_Password") 
$Confirm_Password = $Users_Window.findname("Confirm_Password") 
$Set_Password = $Users_Window.findname("Set_Password") 
$User_Enable_Account = $Users_Window.findname("User_Enable_Account") 


# Block initialization
$Expander_Change_Password_Block.Visibility = "Collapsed"
$Expander_GroupMembership_Block.Visibility = "Collapsed"
$Expander_Selection_Block.Visibility = "Visible"
$Global_Warning_Block.Visibility = "Collapsed"
$User_Block_OK.Visibility = "Collapsed"
$User_Block_KO.Visibility = "Collapsed"


# Expander initialization
$Expander_GroupMembership.IsExpanded = $false
$Expander_Selection.IsExpanded = $True



# ----------------------------------------------------
# Part - Computer GUI
# ----------------------------------------------------

[xml]$XAML_Computers =  
'
<Controls:MetroWindow 
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	xmlns:i="http://schemas.microsoft.com/expression/2010/interactivity"		
	xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
	Title="Quick AD Support - Computer Analysis Part - Part v1.1" 
	Width="470" 
	ResizeMode="NoResize"
	Height="Auto" 
	SizeToContent="Height" 			
	BorderBrush="DodgerBlue"
	BorderThickness="0.5"
	WindowStartupLocation ="CenterScreen"	
	>

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="C:\ProgramData\Quick_AD_Support\QADS_Systray\resources\Icons.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Cobalt.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>

   <Controls:MetroWindow.LeftWindowCommands>
        <Controls:WindowCommands>
            <Button>
                <StackPanel Orientation="Horizontal">
                    <Rectangle Width="15" Height="15" Fill="{Binding RelativeSource={RelativeSource AncestorType=Button}, Path=Foreground}">
                        <Rectangle.OpacityMask>
                            <VisualBrush Stretch="Fill" Visual="{StaticResource appbar_monitor}" />							
                        </Rectangle.OpacityMask>
                    </Rectangle>					
                </StackPanel>
            </Button>				
        </Controls:WindowCommands>	
    </Controls:MetroWindow.LeftWindowCommands>		

    <Grid>	
        <StackPanel Orientation="Vertical" Margin="0,0,0,0">
	
			<StackPanel VerticalAlignment="Center">
				<Image Width="150" Height="70" Source="C:\ProgramData\Quick_AD_Support\QADS_Systray\logo.jpg" ></Image>	
			</StackPanel>	

			<StackPanel Margin="0,0,0,0" >	
			
				<StackPanel Name="Computer_Expander_Selection_Block">							
					<Expander Name="Computer_Expander_Selection" Header="Computer selection"  Margin="0,0,0,0"  IsExpanded="False" Height="auto">  
						<Grid Background="white" >
							<StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Orientation="Horizontal">	
								<Label Margin="0,0,0,0" Content="Type computer name" Foreground="Black" FontSize="18"/>
								<StackPanel  Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">												
									<TextBox ToolTipService.ToolTip="Type the user name you are looking for" Name="Computer_TxtBox" Width="120" FontSize="16"></TextBox>
									<Button Width="40" Name="Check_Computer_btn" BorderThickness="0" Margin="0,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#2196f3">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_magnify}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>	
									
									<StackPanel Name="Computer_Block_OK">
										<Button Width="40" Name="Computer_OK" BorderThickness="0" Margin="5,0,0,0" 
											Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#00a300">
											<Rectangle Width="20" Height="20"  Fill="white" >
												<Rectangle.OpacityMask>
													<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_check}"/>
												</Rectangle.OpacityMask>
											</Rectangle>
										</Button>	
									</StackPanel>
									
									<StackPanel Name="Computer_Block_KO">
										<Button Width="40" x:Name="Computer_KO" BorderThickness="0" Margin="5,0,0,0" 
											Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="Red">
											<Rectangle Width="20" Height="20"  Fill="white" >
												<Rectangle.OpacityMask>
													<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_warning}"/>
												</Rectangle.OpacityMask>
											</Rectangle>
										</Button>	
									</StackPanel>
									
								</StackPanel>				
							</StackPanel>									
						</Grid>
					</Expander> 
				</StackPanel>									
				
				<StackPanel Name="Computer_Global_Warning_Block">
					<Expander Name="Computer_Expander_Warning" Header="Warning infos"  Margin="0,0,0,0" IsExpanded="False" Height="auto">  
						<Grid Background="white">
							<StackPanel Orientation="Vertical">						
								<StackPanel Orientation="Vertical" Margin="0,0,0,0" HorizontalAlignment="Center" VerticalAlignment="Center">				
									<Label Name="Computer_Warning_Label" />
								</StackPanel>					

								<StackPanel Name="DatagGrid_MultiComputers" Margin="0,10,0,0" HorizontalAlignment="Center" VerticalAlignment="Center">
									<Grid Background="CornFlowerBlue">								
								
									<DataGrid IsReadOnly="True" RowHeaderWidth="0" SelectionMode="Single" Height="100" AutoGenerateColumns="False" Name="DatagGrid_Computers" ItemsSource="{Binding}" Margin="1,1,1,1">
										<DataGrid.Columns>										
											<DataGridTextColumn Width="auto" Header="Name" Binding="{Binding Name}"/>		
											<DataGridTextColumn Width="auto" Header="IP" Binding="{Binding IP}"/>							
											<DataGridTextColumn Width="auto" Header="OS" Binding="{Binding OS}"/>												
										</DataGrid.Columns>
									</DataGrid>	
									</Grid>				
									
								</StackPanel>		
							</StackPanel>													
						</Grid>
					</Expander> 				
				</StackPanel>

				<StackPanel Name="Computer_Expander_Basic_Block">				
					<Expander Name="Computer_Expander_Basic" Header="Basic infos"  Margin="0,0,0,0" Background="gray" IsExpanded="False" Height="auto"> 
						<ScrollViewer  CanContentScroll="True" Height="100">        					
							<Grid Background="white" HorizontalAlignment="Left">
								<StackPanel Orientation="Vertical" Margin="0,0,0,0">	
									<StackPanel Orientation="Horizontal">
										<Label Content="Is this computer enabled ?"/>
										<Label Name="Computer_IsEnabled_Value"/>				
									</StackPanel>		

									<StackPanel Orientation="Horizontal">
										<Label Content="Is this computer locked ?"/>
										<Label Name="Computer_Locked_Value"/>				
									</StackPanel>	
																										
									<StackPanel Orientation="Horizontal">
										<Label Content="Is password expired ?"/>
										<Label Name="Computer_IsPWDExpired_Value"/>				
									</StackPanel>		

									<StackPanel Orientation="Horizontal">
										<Label Content="Is this computer reachable ?"/>
										<Label Name="Ping_Status"/>				
									</StackPanel>										
									
									<StackPanel Orientation="Horizontal">
										<Label Content="Name:"/>
										<Label Name="Computer_DisplayName_Value"/>				
									</StackPanel>	
								
									<StackPanel Orientation="Horizontal">
										<Label Content="Password last change:"/>
										<Label Name="Computer_PWDLastChange_Value"/>				
									</StackPanel>										

									<StackPanel Orientation="Horizontal">
										<Label Content="Last logon date:"/>
										<Label Name="Computer_LastLogOn_Value"/>				
									</StackPanel>									

									<StackPanel Orientation="Horizontal">
										<Label Content="Computer OU:"/>
										<TextBox Name="ComputerOU_Value" Width="300"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Computer SID:"/>
										<Label Name="SID_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="DNS Host name:"/>
										<Label Name="DNSName_Value"/>				
									</StackPanel>	
									
									<StackPanel Orientation="Horizontal">
										<Label Content="Canonical name:"/>
										<TextBox Name="CanonicalName_Value" Width="300"/>				
									</StackPanel>										

									<StackPanel Orientation="Horizontal">
										<Label Content="Primary goup:"/>
										<TextBox Name="PrimaryGroup_Value" Width="300"/>				
									</StackPanel>										

								</StackPanel>									
							</Grid>
						</ScrollViewer> 										
					</Expander> 
				</StackPanel>


				<StackPanel Name="Computer_Expander_More_Block">								
					<Expander Name="Computer_Expander_More" Header="More informations" Margin="0,0,0,0" IsExpanded="False" Height="auto">  
						<ScrollViewer HorizontalScrollBarVisibility="Auto" CanContentScroll="True" Height="100">        
							<Grid Background="white" HorizontalAlignment="Left" >
								<StackPanel Orientation="Vertical" Margin="0,0,0,0">						
									<StackPanel Orientation="Horizontal">
										<Label Content="Location:"/>
										<Label Name="Location_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Object category:"/>
										<TextBox Name="ObjCat_Value" Width="300"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Desctiption:"/>
										<Label Name="Description_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="OS Version:"/>
										<Label Name="OSVersion_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="OS:"/>
										<Label Name="OS_Value"/>				
									</StackPanel>										

									<StackPanel Orientation="Horizontal">
										<Label Content="OS Service Pack:"/>
										<Label Name="OSServicePack_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="IP Address:"/>
										<Label Name="IP_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Account created on:"/>
										<Label Name="Computer_WhenCreated_Value"/>				
									</StackPanel>	

									<StackPanel Orientation="Horizontal">
										<Label Content="Account changed on:"/>
										<Label Name="Computer_WhenChanged_Value"/>				
									</StackPanel>										
								</StackPanel>	
							</Grid>							
						</ScrollViewer> 											
					</Expander> 
				</StackPanel>	

				<StackPanel Name="Computer_Expander_More_options_Block">												
					<Expander Name="Computer_Expander_More_options" Header="Other options"  Margin="0,0,0,0" Background="gray" IsExpanded="False" Height="auto">  
						<Grid Background="white" HorizontalAlignment="Center">
							<StackPanel Orientation="Vertical" Margin="0,0,0,0">						
								<StackPanel Orientation="Horizontal">
								
									<Button Width="40" Name="Computer_Export_Values" BorderThickness="0" Margin="0,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#2196f3">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_monitor}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>	
								
								
									<Button Width="40" Name="Computer_Unlock_Account" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#603cba">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_unlock}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>	
									
									<Button Width="40" ToolTip="Enable this account" Name="Computer_Enable_Account" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="black">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_power}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>					

									<Button Width="40" ToolTip="Delete this account" Name="Delete_Computer_Account" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="red">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_delete}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>											

									<Button Width="40" Name="Computer_Display_Groups" BorderThickness="0" Margin="5,0,0,0" 
										Style="{DynamicResource SquareButtonStyle}" Cursor="Hand" Background="#2b5797">
										<Rectangle Width="20" Height="20"  Fill="white" >
											<Rectangle.OpacityMask>
												<VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_group}"/>
											</Rectangle.OpacityMask>
										</Rectangle>
									</Button>																				 
								</StackPanel>
							</StackPanel>						
						</Grid>
					</Expander> 				
				</StackPanel>	
				
				<StackPanel Name="Computer_Expander_GroupMembership_Block">
					<Expander Name="Computer_Expander_GroupMembership" Header="Group membership"  Margin="0,0,0,0" IsExpanded="False" Height="auto">  
						<Grid Background="white">
							<StackPanel Orientation="Vertical">							
								<DataGrid HeadersVisibility="Row" IsReadOnly="True" RowHeaderWidth="0" SelectionMode="Single" Height="100" AutoGenerateColumns="False" Name="Computer_DatagGrid_Groups" ItemsSource="{Binding}" Margin="2,2,2,2" >										
									<DataGrid.Columns>										
										<DataGridTextColumn Width="auto" Header="Name" Binding="{Binding Name}"/>							
									</DataGrid.Columns>
								</DataGrid>	
							</StackPanel>													
						</Grid>
					</Expander> 				
				</StackPanel>					
			</StackPanel>	
        </StackPanel>		
    </Grid>
</Controls:MetroWindow>                
'
$Computers_Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAML_Computers))


# Selection Part
$Computer_Expander_Selection_Block = $Computers_Window.findname("Computer_Expander_Selection_Block") 
$Computer_Expander_Selection = $Computers_Window.findname("Computer_Expander_Selection") 
$Computer_TxtBox = $Computers_Window.findname("Computer_TxtBox") 
$Check_Computer_btn = $Computers_Window.findname("Check_Computer_btn") 
$Computer_Block_OK = $Computers_Window.findname("Computer_Block_OK") 
$Computer_Block_KO = $Computers_Window.findname("Computer_Block_KO") 
$Computer_OK = $Computers_Window.findname("Computer_Expander_More") 
$Computer_KO = $Computers_Window.findname("Computer_Expander_More") 
$Computer_Expander_Selection = $Computers_Window.findname("Computer_Expander_Selection") 



# Warning block
$Computer_Global_Warning_Block = $Computers_Window.findname("Computer_Global_Warning_Block") 
$Computer_Expander_Warning = $Computers_Window.findname("Computer_Expander_Warning") 
$Computer_Warning_Label = $Computers_Window.findname("Computer_Warning_Label") 
$DatagGrid_MultiComputers = $Computers_Window.findname("DatagGrid_MultiComputers") 
$DatagGrid_Computers = $Computers_Window.findname("DatagGrid_Computers") 


# Basic infos block
$Computer_IsEnabled_Value = $Computers_Window.findname("Computer_IsEnabled_Value") 
$Computer_Computer_Expander_Basic_Block = $Computers_Window.findname("Computer_Computer_Expander_Basic_Block") 
$Computer_Expander_Basic = $Computers_Window.findname("Computer_Expander_Basic") 
$Computer_IsPWDExpired_Value = $Computers_Window.findname("Computer_IsPWDExpired_Value") 
$Computer_DisplayName_Value = $Computers_Window.findname("Computer_DisplayName_Value") 
$Computer_PWDLastChange_Value = $Computers_Window.findname("Computer_PWDLastChange_Value") 
$Computer_LastLogOn_Value = $Computers_Window.findname("Computer_LastLogOn_Value") 
$ComputerOU_Value = $Computers_Window.findname("ComputerOU_Value") 
$SID_Value = $Computers_Window.findname("SID_Value") 
$DNSName_Value = $Computers_Window.findname("DNSName_Value") 
$CanonicalName_Value = $Computers_Window.findname("CanonicalName_Value") 
$PrimaryGroup_Value = $Computers_Window.findname("PrimaryGroup_Value") 
$Computer_Locked_Value = $Computers_Window.findname("Computer_Locked_Value") 
$Ping_Status = $Computers_Window.findname("Ping_Status") 


# More infos block
$Computer_Computer_Expander_More_Block = $Computers_Window.findname("Computer_Computer_Expander_More_Block") 
$Computer_Expander_More = $Computers_Window.findname("Computer_Expander_More") 
$Location_Value = $Computers_Window.findname("Location_Value") 
$ObjCat_Value = $Computers_Window.findname("ObjCat_Value") 
$Description_Value = $Computers_Window.findname("Description_Value") 
$OSVersion_Value = $Computers_Window.findname("OSVersion_Value") 
$OS_Value = $Computers_Window.findname("OS_Value") 
$OSServicePack_Value = $Computers_Window.findname("OSServicePack_Value") 
$IP_Value = $Computers_Window.findname("IP_Value") 


$Computer_WhenCreated_Value = $Computers_Window.findname("Computer_WhenCreated_Value") 
$Computer_WhenChanged_Value = $Computers_Window.findname("Computer_WhenChanged_Value") 


# More options action block
$Computer_Expander_More_options_Block = $Computers_Window.findname("Computer_Expander_More_options_Block") 
$Computer_Export_Values = $Computers_Window.findname("Computer_Export_Values") 
$Computer_Unlock_Account = $Computers_Window.findname("Computer_Unlock_Account") 
$Computer_Display_Groups = $Computers_Window.findname("Computer_Display_Groups") 
$Computer_Enable_Account = $Computers_Window.findname("Computer_Enable_Account") 
$Delete_Computer_Account = $Computers_Window.findname("Delete_Computer_Account") 



# Group membership block
$Computer_Expander_GroupMembership_Block = $Computers_Window.findname("Computer_Expander_GroupMembership_Block") 
$Computer_Expander_GroupMembership = $Computers_Window.findname("Computer_Expander_GroupMembership") 
$Computer_DatagGrid_Groups = $Computers_Window.findname("Computer_DatagGrid_Groups") 

# Block initialization
$Computer_Expander_GroupMembership_Block.Visibility = "Collapsed"
$Computer_Global_Warning_Block.Visibility = "Collapsed"
$Computer_Block_OK.Visibility = "Collapsed"
$Computer_Block_KO.Visibility = "Collapsed"
$Computer_Expander_Selection_Block.Visibility = "Visible"

# Expander initialization
$Computer_Expander_GroupMembership.IsExpanded = $false
$Computer_Expander_Selection.IsExpanded = $True












# ----------------------------------------------------
# Part - Computer GUI
# ----------------------------------------------------

[xml]$XAML_Admin =  
'
<Controls:MetroWindow 
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	xmlns:i="http://schemas.microsoft.com/expression/2010/interactivity"		
	xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
	Title="Quick AD Support - Admin Part" 
	Width="455" 
	Height="280" 
	ResizeMode="NoResize"		
	BorderBrush="DodgerBlue"
	BorderThickness="0.5"
	WindowStartupLocation ="CenterScreen"	
	>

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="C:\ProgramData\Quick_AD_Support\QADS_Systray\resources\Icons.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Cobalt.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>

   <Controls:MetroWindow.LeftWindowCommands>
        <Controls:WindowCommands>
            <Button>
                <StackPanel Orientation="Horizontal">
                    <Rectangle Width="15" Height="15" Fill="{Binding RelativeSource={RelativeSource AncestorType=Button}, Path=Foreground}">
                        <Rectangle.OpacityMask>
                            <VisualBrush Stretch="Fill" Visual="{StaticResource appbar_lock}" />							
                        </Rectangle.OpacityMask>
                    </Rectangle>					
                </StackPanel>
            </Button>				
        </Controls:WindowCommands>	
    </Controls:MetroWindow.LeftWindowCommands>		

    <Grid>	
		<StackPanel  HorizontalAlignment="Center" VerticalAlignment="Center">



			<StackPanel Orientation="Horizontal">
				<StackPanel>
					<Rectangle  Width="70" Height="70"  Fill="gray" Margin="0,7,0,0">
						<Rectangle.OpacityMask>
							<VisualBrush Stretch="Fill" Visual="{StaticResource appbar_information_circle}"/>
						</Rectangle.OpacityMask>
					</Rectangle>
				</StackPanel>				
			
				<StackPanel Margin="10,0,0,0">
					<Border BorderBrush="DodgerBlue" BorderThickness="1" Width="305" Height="80">			
						<StackPanel Orientation="Vertical"  Margin="10,10,10,10">						
							<StackPanel Orientation="Vertical"  Margin="0,0,0,0" VerticalAlignment="Center">
								<Label Content="Working with another account" FontWeight="Bold" Foreground="Blue" HorizontalAlignment="Center"/>
								<Label Name="Admin_Account_Status"  Foreground="Blue" HorizontalAlignment="Center"/>								
							</StackPanel>					
						</StackPanel>	
					</Border>	
				</StackPanel>		
				
			</StackPanel>	

		
			<StackPanel Orientation="Horizontal" Margin="0,5,0,0">
				<StackPanel Name="Lock">
					<Rectangle  Width="70" Height="90"  Fill="Gray" Margin="0,10,0,0">
						<Rectangle.OpacityMask>
							<VisualBrush Stretch="Fill" Visual="{StaticResource appbar_lock}"/>
						</Rectangle.OpacityMask>
					</Rectangle>
				</StackPanel>	
				
			
				<StackPanel Margin="10,0,0,0">
					<Border BorderBrush="DodgerBlue" BorderThickness="1" Width="305">			
						<StackPanel Orientation="Vertical"  Margin="10,10,10,10">
						
							<StackPanel Orientation="Horizontal"  Margin="0,0,0,0" >
								<Label Content="Type your user name" Width="130"/>
								<TextBox Name="Admin_User"  Width="150" />
							</StackPanel>	
							
							<StackPanel Orientation="Horizontal"  Margin="0,10,0,0" >
								<Label Content="Type your password" Width="130"/>								
								<PasswordBox  
								Name="Admin_Password"  Width="150"  
								Controls:TextBoxHelper.ClearTextButton="{Binding RelativeSource={RelativeSource Self}, Path=(Controls:TextBoxHelper.HasText), Mode=OneWay}" 
								Controls:TextBoxHelper.IsWaitingForData="True" 
								Controls:TextBoxHelper.Watermark="Local Admin Password" 	
								Style="{StaticResource MetroButtonRevealedPasswordBox}"								
								/>								

							</StackPanel>		

							<StackPanel Orientation="Horizontal"  Margin="0,10,0,0" >
								<Label Content="" Width="130"/>					
								<Button Name="Set_Admin_Account" Content="Set this account" Width="150" Background="#00a300" Foreground="White" BorderThickness="0"/>
							</StackPanel>					
						</StackPanel>	
					</Border>	
				</StackPanel>						
			</StackPanel>		
		</StackPanel>				
    </Grid>
</Controls:MetroWindow>                
'
$Admin_Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAML_Admin))


# Selection Part
$Admin_User = $Admin_Window.findname("Admin_User") 
$Admin_Password = $Admin_Window.findname("Admin_Password") 
$Set_Admin_Account = $Admin_Window.findname("Set_Admin_Account") 
$Lock = $Admin_Window.findname("Lock") 
$Admin_Account_Status = $Admin_Window.findname("Admin_Account_Status") 



# ----------------------------------------------------
# Part - About
# ----------------------------------------------------

[xml]$xaml_About =  
'<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
WindowStyle="None" 
Height="260" 
Width="380"
ResizeMode="NoResize" 
ShowInTaskbar="False"
Background="Transparent"
AllowsTransparency="True" 
>

<Border  BorderBrush="Black" BorderThickness="1" Margin="10,10,10,10">
<Grid Name="grid" Background="White" >		
	<StackPanel HorizontalAlignment="Center"  VerticalAlignment="Center" Orientation="Vertical">
		<StackPanel Height="200"  HorizontalAlignment="Center"  VerticalAlignment="Top">		
			<Image Width="90" Height="90" Source="C:\ProgramData\Quick_AD_Support\QADS_Systray\logo.png" Margin="0,30,0,0"/>					
			<Label Margin="0,0,0,0" x:Name="About_Version" Content="Version: 1.0" Foreground="Black" FontSize="12" HorizontalAlignment="Center"/>
			<Label Margin="0,-5,0,0" x:Name="About_Date" Content="Last release: 20/09/19" Foreground="Black" FontSize="12" HorizontalAlignment="Center"/>				
			<Label Margin="0,-5,0,0" x:Name="About_Name" Content="Damirn van robaeys" Foreground="Black" FontSize="12"  HorizontalAlignment="Center"/>
		</StackPanel>
		
		<StackPanel Height="40" Name="Main_Update_Status" Background="#F2F2F2" Width="370" HorizontalAlignment="Center" VerticalAlignment="Bottom">	
				<Line Grid.Row="1" X1="0" Y1="0" X2="1"  Y2="0" Stroke="Black" StrokeThickness="0.2" Stretch="Uniform"></Line>						
				<Label Margin="0,10,0,0" Name="Update_Status_Label" Content="The tool is up to date" FontSize="10" HorizontalAlignment="Center" VerticalAlignment="Center"/>
		</StackPanel>		
	</StackPanel>		
</Grid>	
</Border>
</Window>'

$window_About = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml_About))
$About_Version = $window_About.findname("About_Version") 
$About_Date = $window_About.findname("About_Date") 
$About_Name = $window_About.findname("About_Name") 
$Main_Update_Status = $window_About.findname("Main_Update_Status") 
$Update_Status_Label = $window_About.findname("Update_Status_Label") 

$About_Version.Content = "Version: 1.1"
$About_Date.Content = "Release date: 11/13/2018"
$About_Name.Content = "Author: Damien Van Robaeys"




########################################################################################################################################################################################################	
#*******************************************************************************************************************************************************************************************************
#																						MAIN FUNCTIONS
#*******************************************************************************************************************************************************************************************************
########################################################################################################################################################################################################


# ----------------------------------------------------------------------
# Part - Function for the warning part
# ----------------------------------------------------------------------	

# Warning initialization
Function Warning_Users_Block
	{
		$Warning_Label.Foreground = "red"
		$Warning_Label.FontSize = "16"
	
		$Global_Warning_Block.Visibility = "Visible"					
	
		$Expander_Warning.IsExpanded = $True										
	
		$User_Block_OK.Visibility = "Collapsed"
		$User_Block_KO.Visibility = "Visible"
		
		$Expander_Selection_Block.Visibility = "Visible"	
		
		$Expander_Basic.IsExpanded = $False
		$Expander_GroupMembership_Block.Visibility = "Collapsed"
		$Expander_GroupMembership.IsExpanded = $False
	}
	
Function Disable_Warning_Users_Block
	{
		$Warning_Label.Foreground = ""
		$Warning_Label.FontSize = ""
	
		$Global_Warning_Block.Visibility = "Collapsed"					
	
		$Expander_Warning.IsExpanded = $false										
	
		$User_Block_OK.Visibility = "Visible"
		$User_Block_KO.Visibility = "Collapsed"
		
		$Expander_Selection_Block.Visibility = "Visible"	
		
		$Expander_Basic.IsExpanded = $true
		$Expander_GroupMembership_Block.Visibility = "Collapsed"
		$Expander_GroupMembership.IsExpanded = $False
	}	
	
	
Function Warning_Computers_Block
	{
		$Computer_Warning_Label.Foreground = "red"
		$Computer_Warning_Label.FontSize = "16"
	
		$Computer_Global_Warning_Block.Visibility = "Visible"					
	
		$Computer_Expander_Warning.IsExpanded = $True										
	
		$Computer_Block_OK.Visibility = "Collapsed"
		$Computer_Block_KO.Visibility = "Visible"
		
		$Computer_Expander_Selection_Block.Visibility = "Visible"	
		
		$Computer_Expander_Basic.IsExpanded = $False
		$Computer_Expander_GroupMembership_Block.Visibility = "Collapsed"
		$Computer_Expander_GroupMembership.IsExpanded = $False		
		
		$DatagGrid_MultiComputers.Visibility = "Collapsed"
	}	

	
	
# ----------------------------------------------------------------------
# Part - Function for displaying infos when a user has been found
# ----------------------------------------------------------------------		

	
Function User_Found_Custo
	{
		$User_Block_OK.Visibility = "Visible"
		$User_Block_KO.Visibility = "Collapsed"		
		
		$DatagGrid_MultiUsers.Visibility = "Collapsed"
			
		$Global_Warning_Block.Visibility = "Collapsed"					
	
		$Expander_Warning.IsExpanded = $False										
	
		$User_Block_OK.Visibility = "Visible"
		$User_Block_KO.Visibility = "Collapsed"
		
		$Expander_Selection_Block.Visibility = "Visible"
		$Expander_Basic_Block.Visibility = "Visible"
		$Expander_More_Block.Visibility = "Visible"
		$Expander_More_options_Block.Visibility = "Visible"									

		$Account_Value.Content = $Get_User_Infos.SamAccountName
		$DisplayName_Value.Content = $Get_User_Infos.Name
		$UserOU_Value.Text = $Get_User_Infos.DistinguishedName	
		$UserOU_Value.BorderBrush = "Transparent"
		$LastLogOn_Value.Content = $Get_User_Infos.LastLogonDate	
		$Locked_Value.Content = $Get_User_Infos.LockedOut		
		
		$Dept_Value.Content = $Get_User_Infos.Department	 
		$IsEnabled_Value.Content = $Get_User_Infos.Enabled			
		$LastBadPWD_Value.Content = $Get_User_Infos.LastBadPasswordAttempt									
		$Mail_Value.Content = $Get_User_Infos.Mail	
		$Office_Value.Content = $Get_User_Infos.Office	
		$IsPWDExpired_Value.Content = $Get_User_Infos.PasswordExpired									
		$PWDLastChange_Value.Content = $Get_User_Infos.PasswordLastSet		
		$PWDNeverExpires_Value.Content = $Get_User_Infos.PasswordNeverExpires									
		$WhenChanged_Value.Content = $Get_User_Infos.WhenChanged									
		$WhenCreated_Value.Content = $Get_User_Infos.WhenCreated	
		$CannotChangePassword_Value.Content = $Get_User_Infos.CannotChangePassword	

		
		$User = $Get_User_Infos.SamAccountName						
		$PWD_Expire_Date = (Get-ADUser -filter {((SamAccountName -like $User))} -Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed" |
		Select-Object -Property "Displayname",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}).ExpiryDate		

		$PWD_Expiration_Date_Value.Content = $PWD_Expire_Date

		$Get_User_Groups = (Get-ADPrincipalGroupMembership ($Get_User_Infos.SamAccountName))
		ForEach ($Group in $Get_User_Groups)				
			{
				$Groups_values = New-Object PSObject
				$Groups_values = $Groups_values | Add-Member Name $Group.name -passthru		
				$DatagGrid_Groups.Items.Add($Groups_values) > $null		
			}											

		If (($Get_User_Infos.LockedOut) -eq $true)
			{
				$Locked_Value.Foreground = "Red"
				Warning_Users_Block								
				$Warning_Label.Content = "Account $MyUser is locked"	
				$User_Unlock_Account.IsEnabled = $True	

				$Expander_More_options.IsExpanded = $True				
				$Expander_Basic.IsExpanded = $False
				$User_Unlock_Account.BorderBrush = "Red"	
				$User_Unlock_Account.BorderThickness = "2"
				$User_Unlock_Account.ToolTip = "Unlock the user account"				
			}
		Else
			{
				$Locked_Value.Foreground = "Blue"
				$User_Unlock_Account.IsEnabled = $true		
				$User_Unlock_Account.ToolTip = "This user account is not locked"								
			}									
		
		If (($Get_User_Infos.Enabled) -eq $false) 
			{
				$IsEnabled_Value.Foreground = "Red"
				Warning_Users_Block									
				$Warning_Label.Content = "Account $MyUser is disabled"					
				$Expander_More_options.IsExpanded = $True				
				$Expander_Basic.IsExpanded = $False
				$User_Enable_Account.IsEnabled = $true
				$User_Enable_Account.BorderBrush = "Red"	
				$User_Enable_Account.BorderThickness = "2"	
				$User_Enable_Account.ToolTip = "Enable the user account"				
				
			}
		Else
			{
				$IsEnabled_Value.Foreground = "Blue"
				$User_Enable_Account.IsEnabled = $true	
				$User_Enable_Account.ToolTip = "Disable this user account"								
			}											

		If (($Get_User_Infos.PasswordExpired) -eq $true)  
			{
				Warning_Users_Block
				$IsPWDExpired_Value.Foreground = "Red"
				$Expander_More_options.IsExpanded = $True				
				$Expander_Basic.IsExpanded = $False
				$Change_Password.BorderBrush = "Red"	
				$Change_Password.BorderThickness = "2"	
				$Warning_Label.Content = "Password for $MyUser has expired"																						
			}
		Else
			{
				$IsPWDExpired_Value.Foreground = "Blue"
			}	

			
		If ((($Get_User_Infos.LockedOut) -eq $false) -and (($Get_User_Infos.Enabled) -eq $true) -and (($Get_User_Infos.PasswordExpired) -eq $False))
			{
				$Expander_Basic.IsExpanded = $True
				$Expander_More.IsExpanded = $False				
			}	
	}
	
	
# ----------------------------------------------------------------------
# Part - Function for displaying infos when a computer has been found
# ----------------------------------------------------------------------	
	
Function Computer_Found_Custo
	{
		$Computer_Expander_Basic.IsExpanded = $True
		$Computer_Expander_More.IsExpanded = $False		

		$Computer_Block_OK.Visibility = "Visible"
		$Computer_Block_KO.Visibility = "Collapsed"		
		
		$Computer_Global_Warning_Block.Visibility = "Collapsed"						
		$Computer_Expander_Warning.IsExpanded = $False										
	
		$Computer_Expander_Selection_Block.Visibility = "Visible"
		$Computer_Expander_Basic_Block.Visibility = "Visible"
		$Computer_Expander_More_Block.Visibility = "Visible"
		$Computer_Expander_More_options_Block.Visibility = "Visible"									

		$Computer_DisplayName_Value.Content = $Get_Computer_Infos.Name
		$Computer_LastLogOn_Value.Content = $Get_Computer_Infos.LastLogonDate	
		
		$ComputerOU_Value.Text = $Get_Computer_Infos.DistinguishedName	
		$ComputerOU_Value.BorderBrush = "Transparent"	

		$SID_Value.Content = $Get_Computer_Infos.SID	
		$OSVersion_Value.Content = $Get_Computer_Infos.OperatingSystemVersion	
		$OS_Value.Content = $Get_Computer_Infos.OperatingSystem	
		$OSServicePack_Value.Content = $Get_Computer_Infos.OperatingSystemServicePack	
		$IP_Value.Content = $Get_Computer_Infos.IPv4Address	
		$DNSName_Value.Content = $Get_Computer_Infos.DNSHostName
		
		$Computer_IsEnabled_Value.Content = $Get_Computer_Infos.Enabled			
		$Computer_Locked_Value.Content = $Get_Computer_Infos.LockedOut					
		
		$Computer_IsPWDExpired_Value.Content = $Get_Computer_Infos.PasswordExpired									
		$Computer_PWDLastChange_Value.Content = $Get_Computer_Infos.PasswordLastSet	

		$Location_Value.Content = $Get_Computer_Infos.Location			
		$Description_Value.Content = $Get_Computer_Infos.Description			
		$Computer_WhenChanged_Value.Content = $Get_Computer_Infos.WhenChanged									
		$Computer_WhenCreated_Value.Content = $Get_Computer_Infos.WhenCreated	

		$ObjCat_Value.Text = $Get_Computer_Infos.ObjectCategory	
		$ObjCat_Value.BorderBrush = "Transparent"									

		$CanonicalName_Value.Text = $Get_Computer_Infos.CanonicalName
		$CanonicalName_Value.BorderBrush = "Transparent"

		$PrimaryGroup_Value.Text = $Get_Computer_Infos.PrimaryGroup	
		$PrimaryGroup_Value.BorderBrush = "Transparent"	
								
		$Get_Computer_Group = Get-ADPrincipalGroupMembership (Get-ADComputer $Get_Computer_Infos.Name).DistinguishedName
		ForEach ($Group in $Get_Computer_Group)				
			{
				$Groups_values = New-Object PSObject
				$Groups_values = $Groups_values | Add-Member Name $Group.name -passthru		
				$Computer_DatagGrid_Groups.Items.Add($Groups_values) > $null		
			}											

		If (($Get_Computer_Infos.Enabled) -eq $false) 
			{
				$Computer_IsEnabled_Value.Foreground = "Red"
				Warning_Computers_Block									
				$Computer_Warning_Label.Content = "$MyComputer is disabled"					
				$Computer_Expander_More_options.IsExpanded = $True				
				$Computer_Expander_Basic.IsExpanded = $False
				$Computer_Enable_Account.IsEnabled = $true
				$Computer_Enable_Account.BorderBrush = "Red"	
				$Computer_Enable_Account.BorderThickness = "2"	
				$Computer_Enable_Account.ToolTip = "Enable this computer"				
				
			}
		Else
			{
				$Computer_IsEnabled_Value.Foreground = "Blue"
				$Computer_Enable_Account.IsEnabled = $true	
				$Computer_Enable_Account.ToolTip = "Disable the computer"								
			}		

			
			
		If (($Get_Computer_Infos.LockedOut) -eq $true)
			{
				$Computer_Locked_Value.Foreground = "Red"
				Warning_Computers_Block								
				$Computer_Warning_Label.Content = "$MyComputer is locked"	
				$Computer_Unlock_Account.IsEnabled = $True	

				$Computer_Expander_More_options.IsExpanded = $True				
				$Computer_Expander_Basic.IsExpanded = $False
				$Computer_Unlock_Account.BorderBrush = "Red"	
				$Computer_Unlock_Account.BorderThickness = "2"
				$Computer_Unlock_Account.ToolTip = "Unlock this computer"				
			}
		Else
			{
				$Computer_Locked_Value.Foreground = "Blue"
				$Computer_Unlock_Account.IsEnabled = $False		
				$Computer_Unlock_Account.ToolTip = "This computer is not locked"								
			}		


		If (($Get_Computer_Infos.PasswordExpired) -eq $true)  
			{
				Warning_Computers_Block
				$Computer_IsPWDExpired_Value.Foreground = "Red"
				$Computer_Expander_More_options.IsExpanded = $True				
				$Computer_Expander_Basic.IsExpanded = $False
				$Warning_Label.Content = "Password for $MyUser has expired"																						
			}
		Else
			{
				$Computer_IsPWDExpired_Value.Foreground = "Blue"
			}

		
		$Comp_Name = $Get_Computer_Infos.Name		
		$Test_connection = Test-Connection $Comp_Name -count 1 -quiet
		If ($Test_connection)
			{	
				$Ping_Status.Content = "True"
				$Ping_Status.Foreground = "Blue"				
			}
		Else
			{
				$Ping_Status.Content = "False"
				$Ping_Status.Foreground = "Red"							
			}
	}	
	
	
	
# ----------------------------------------------------------------------
# Part - Function to clear user informations
# ----------------------------------------------------------------------		

Function Clear_Users_Datas
	{
		$User_TxtBox.Text = ""
		
		$User_Block_OK.Visibility = "Collapsed"
		$User_Block_KO.Visibility = "Collapsed"			
		$Global_Warning_Block.Visibility = "Collapsed"						
	
		$User_Block_OK.Visibility = "Collapsed"
		$User_Block_KO.Visibility = "Collapsed"
		$Expander_GroupMembership_Block.Visibility = "Collapsed"
		
		$Expander_Selection_Block.Visibility = "Visible"
		$Expander_Basic_Block.Visibility = "Visible"
		$Expander_More_Block.Visibility = "Visible"
		$Expander_More_options_Block.Visibility = "Visible"		

		$Expander_Change_Password_Block.Visibility = "Collapsed"
		
		$User_Enable_Account.BorderBrush = "Transparent"	
		$User_Enable_Account.BorderThickness = "0"			

		$Account_Value.Content = ""
		$DisplayName_Value.Content = ""
		$UserOU_Value.Text = ""
		$UserOU_Value.BorderBrush = "Transparent"
		$LastLogOn_Value.Content = ""	
		$Locked_Value.Content = ""		
		
		$Dept_Value.Content = ""	 
		$IsEnabled_Value.Content = ""		
		$LastBadPWD_Value.Content = ""								
		$Mail_Value.Content = ""	
		$Office_Value.Content = ""	
		$IsPWDExpired_Value.Content = ""									
		$PWDLastChange_Value.Content = ""	
		$PWDNeverExpires_Value.Content = ""								
		$WhenChanged_Value.Content = ""									
		$WhenCreated_Value.Content = ""
		$CannotChangePassword_Value.Content = ""
											
		$PWD_Expiration_Date_Value.Content = ""
		
		$Get_User_Groups = $null
		$Warning_Label.Content = ""	

		$Expander_More_options.IsExpanded = $False				
		$Expander_Basic.IsExpanded = $False
		$Expander_More.IsExpanded = $False	

	}




# ----------------------------------------------------------------------
# Part - Function to clear computer informations
# ----------------------------------------------------------------------		

Function Clear_Computer_Datas
	{
		$Computer_TxtBox.Text = ""
		
		$Computer_Block_OK.Visibility = "Collapsed"
		$Computer_Block_KO.Visibility = "Collapsed"			
		$Computer_Global_Warning_Block.Visibility = "Collapsed"						
		$Computer_Expander_Warning.IsExpanded = $False		

		$Computer_Expander_Selection.IsExpanded = $True	
	
		$Computer_Block_OK.Visibility = "Collapsed"
		$Computer_Block_KO.Visibility = "Collapsed"
		
		$Computer_Enable_Account.BorderBrush = ""	
		$Computer_Enable_Account.BorderThickness = "0"			
		
		$Computer_Expander_Selection_Block.Visibility = "Visible"
		$Computer_Expander_Basic_Block.Visibility = "Visible"
		$Computer_Expander_More_Block.Visibility = "Visible"
		$Computer_Expander_More_options_Block.Visibility = "Visible"									

		$Computer_IsEnabled_Value.Content = ""
		$Computer_Locked_Value.Content = ""
		$Computer_IsPWDExpired_Value.Content = ""
		$Computer_DisplayName_Value.Content = ""
		$Computer_PWDLastChange_Value.Content = ""
		$Computer_LastLogOn_Value.Content = ""
		$Computer_LastLogOn_Value.Content = ""
		
		$Ping_Status.Content = ""
		
		$ComputerOU_Value.Text = ""
		$ComputerOU_Value.BorderBrush = "Transparent"
		
		$SID_Value.Content = ""
		$DNSName_Value.Content = ""
		
		$CanonicalName_Value.Text = ""
		$CanonicalName_Value.BorderBrush = "Transparent"		
		
		$PrimaryGroup_Value.Content = ""	
		$PrimaryGroup_Value.Content = ""	

		$Location_Value.Content = ""
		
		$ObjCat_Value.Text = ""
		$ObjCat_Value.BorderBrush = "Transparent"	
		
		$Description_Value.Content = ""
		$OSVersion_Value.Content = ""
		$OS_Value.Content = ""
		$OSServicePack_Value.Content = ""
		$IP_Value.Content = ""
		$Computer_WhenCreated_Value.Content = ""
		$Computer_WhenChanged_Value.Content = ""
													
		$Get_Computer_Groups = $null
		$Computer_DatagGrid_Groups.Item.Clear()
		$Computer_Warning_Label.Content = ""	

		$Computer_Expander_More.IsExpanded = $False				
		$Computer_Expander_Basic.IsExpanded = $False

	}





########################################################################################################################################################################################################	
#*******************************************************************************************************************************************************************************************************
#																						 PROGRESSBAR DESIGN USER
#*******************************************************************************************************************************************************************************************************
########################################################################################################################################################################################################

$syncProgress = [hashtable]::Synchronized(@{})
$childRunspace = [runspacefactory]::CreateRunspace()
$childRunspace.ApartmentState = "STA"
$childRunspace.ThreadOptions = "ReuseThread"         
$childRunspace.Open()
$childRunspace.SessionStateProxy.SetVariable("syncProgress",$syncProgress)          
$PsChildCmd = [PowerShell]::Create().AddScript({   
    [xml]$xaml = @"
	<Controls:MetroWindow 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
		xmlns:i="http://schemas.microsoft.com/expression/2010/interactivity"				
		xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"		
        Name="WindowProgress" 
		WindowStyle="None" 
		AllowsTransparency="True" 
		UseNoneWindowStyle="True"	
		Width="470" 
		Height="470" 
		WindowStartupLocation ="CenterScreen"		
		>

	<Window.Resources>
		<ResourceDictionary>
			<ResourceDictionary.MergedDictionaries>
				<ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
				<ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
				<ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
				<ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Cobalt.xaml" />
				<ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
			</ResourceDictionary.MergedDictionaries>
		</ResourceDictionary>
	</Window.Resources>		

	<Window.Background>
		<SolidColorBrush Opacity="0.7" Color="#0077D6"/>
	</Window.Background>	
		
	<Grid HorizontalAlignment="Center" VerticalAlignment="Center">	
		<StackPanel Orientation="Vertical" HorizontalAlignment="Center" Margin="0,0,0,0">	
			<StackPanel Orientation="Vertical" HorizontalAlignment="Center" Margin="0,0,0,0">	
				<Controls:ProgressRing IsActive="True" Margin="0,0,0,0"  Foreground="White" Width="50"/>
			</StackPanel>								
			
			<StackPanel Orientation="Vertical" HorizontalAlignment="Center" Margin="0,0,0,0">				
				<Label Name="ProgressStep" Content="Quick AD Support is looking for your user" FontSize="17" Margin="0,0,0,0" Foreground="White"/>	
			</StackPanel>			
		</StackPanel>										
	</Grid>
</Controls:MetroWindow>
"@
  
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $syncProgress.Window=[Windows.Markup.XamlReader]::Load( $reader )
    $syncProgress.Label = $syncProgress.window.FindName("ProgressStep")	

    $syncProgress.Window.ShowDialog() #| Out-Null
    $syncProgress.Error = $Error
})




################ Launch Progress Bar  ########################  
Function Launch_modal_progress{    
    $PsChildCmd.Runspace = $childRunspace
    $Script:Childproc = $PsChildCmd.BeginInvoke()
	
}

################ Close Progress Bar  ########################  
Function Close_modal_progress{
    $syncProgress.Window.Dispatcher.Invoke([action]{$syncProgress.Window.close()})
    $PsChildCmd.EndInvoke($Script:Childproc) | Out-Null
}




########################################################################################################################################################################################################	
#*******************************************************************************************************************************************************************************************************
#																						 MAIN CONTROLS ACTIONS
#*******************************************************************************************************************************************************************************************************
########################################################################################################################################################################################################




################################################################################################################################"
# ACTIONS FROM THE SYSTRAY
################################################################################################################################"


# ----------------------------------------------------
# Part - Add the systray menu
# ----------------------------------------------------		
	
# Create notifyicon, and right-click -> Exit menu
$Main_Tool_Icon = New-Object System.Windows.Forms.NotifyIcon
$Main_Tool_Icon.Text = "Quick AD Support"
$Main_Tool_Icon.Icon = $icon
$Main_Tool_Icon.Visible = $true

$Menu_Users = New-Object System.Windows.Forms.MenuItem
$Menu_Users.Text = "User analysis"

$Menu_Computers = New-Object System.Windows.Forms.MenuItem
$Menu_Computers.Text = "Computer analysis"

$Menu_Admin = New-Object System.Windows.Forms.MenuItem
$Menu_Admin.Text = "Set admin account"

$Menu_About = New-Object System.Windows.Forms.MenuItem
$Menu_About.Text = "About"

$Menu_Restart_Tool = New-Object System.Windows.Forms.MenuItem
$Menu_Restart_Tool.Text = "Restart the tool (in 10secs)"

$Menu_Exit = New-Object System.Windows.Forms.MenuItem
$Menu_Exit.Text = "Exit"

$contextmenu = New-Object System.Windows.Forms.ContextMenu
$Main_Tool_Icon.ContextMenu = $contextmenu
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Users)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Computers)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Admin)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_About)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Restart_Tool)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Exit)


# ---------------------------------------------------------------------
# Action when after a click on the systray icon
# ---------------------------------------------------------------------
$Main_Tool_Icon.Add_Click({
	Clear_Users_Datas
					
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Users_Window)

	If ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
		$Users_Window.WindowStartupLocation = "CenterScreen"	
		$Users_Window.Show()
		$Users_Window.Activate()	
	}				
})



# ---------------------------------------------------------------------
# Action after clicking on User Analysis
# ---------------------------------------------------------------------
$Menu_Users.Add_Click({
	Clear_Users_Datas					
	
	$Users_Window.WindowStartupLocation = "CenterScreen"	
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Users_Window)
	$Users_Window.ShowDialog()
	$Users_Window.Activate()	
	
	
})


# ---------------------------------------------------------------------
# Action after clicking on Computer Analysis
# ---------------------------------------------------------------------
$Menu_Computers.Add_Click({
	Clear_Computer_Datas
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Computers_Window)
	$Computers_Window.Show()
	$Computers_Window.Activate()	
})





# ---------------------------------------------------------------------
# Action after clicking on Admin part
# ---------------------------------------------------------------------
$Menu_Admin.Add_Click({
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Admin_Window)
	$Admin_Window.Show()
	$Admin_Window.Activate()	
	$admin_user.Text = ""						
	$admin_password.password = ""
	$Lock.Fill = 'Transparent'		

	
	If ((test-path $File_User) -and (test-path $file))
		{
			# $Admin_Account_Status.Content = "An admin account has already been setted"			
			$admin_user.Text = get-content $File_User		
			$admin_password.password = ""									
		}
	Else
		{
			If (test-path $File_User)
				{
					$admin_user.Text = ""	
				}
				
			$Admin_Account_Status.Content = "Type your admin account credentials"	
		}
})





############################################################################################################
# 								ACTION ON CLICK ON THE CLOSE BUTTON
############################################################################################################

# ---------------------------------------------------------------------
# Action on close user GUI
# ---------------------------------------------------------------------

# Action on the close button
$Users_Window.Add_Closing({
	$_.Cancel = $true
	[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Oops :-(", "To close the window click out of the window !!!")					
})

# ---------------------------------------------------------------------
# Action on close computer GUI
# ---------------------------------------------------------------------

$Computers_Window.Add_Closing({
	$_.Cancel = $true
	[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Oops :-(", "To close the window click out of the window !!!")					
})


# ---------------------------------------------------------------------
# Action on close Admin part GUI
# ---------------------------------------------------------------------

$Admin_Window.Add_Closing({
	$_.Cancel = $true
	[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Admin_Window, "Oops :-(", "To close the window click out of the window !!!")					
})


$Admin_Window.Add_Deactivated({
	$Admin_Window.Hide()	
})



$Use_Creds_Status = $false
$ProgData = $env:PROGRAMDATA
$QADS_Tool = "$ProgData\Quick_AD_Support\Infos"
$File = "$QADS_Tool\QADS_SD.txt"	
$File_User = "$QADS_Tool\QADS_User.txt"

############################################################################################################
# 								ACTION ON CLICK ON THE SET THIS ACCOUNT BUTTON
############################################################################################################

# $Script:Test_Admin_Account = "False"	

If ((test-path $File_User) -and (test-path $file))
	{
		$Admin_Account_Status.Content = "An admin account has already been setted"			
	}
Else
	{
		$Admin_Account_Status.Content = "Type your admin account credentials"	
	}
		

$Set_Admin_Account.Add_Click({
	$User_Name = $admin_user.Text.ToString()
	$User_Password = $admin_password.Password
	
	# If there is no user name in the textbox
	If (($admin_user.Text) -eq "")
		{
			$admin_user.BorderBrush = "Red"
			$User_Name.BorderThickness = "1"			
		}
		
	# If there is no password in the passwordbox		
	If (($admin_password.Password) -eq "")
		{
			$admin_password.BorderBrush = "Red"
			$User_Name.BorderThickness = "1"			
		}		

	$Global:Admin_User_Name = $User_Name
	[Byte[]] $Global:key = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
	
	$ProgData = $env:PROGRAMDATA
	$QADS_Tool_Infos = "$ProgData\Quick_AD_Support\Infos"
	$File = "$QADS_Tool_Infos\QADS_SD.txt"
	
	$File_User = "$QADS_Tool_Infos\QADS_User.txt"
	If (test-path $File_User)
		{
			remove-item $File_User -force
		}
	new-item $File_User -type file					
	Add-content $File_User $User_Name 


	$Current_User = $env:username
	$Password = $User_Password | ConvertTo-SecureString -AsPlainText -Force
	$Password | ConvertFrom-SecureString -key $key | Out-File $File		
	# $Global:Encrypted_Password = $Password | ConvertFrom-SecureString -key $key 				
	$Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, (Get-Content $File | ConvertTo-SecureString -Key $key)		
	# $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, ($Encrypted_Password | ConvertTo-SecureString -Key $key)		
	
	
	$Check_Admin_User = get-aduser $Admin_User_Name 
	If ($Check_Admin_User -eq $null)
		{
			[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Admin_Window, "Success :-)", "can not found user $Admin_User_Name")			
		}
	Else
		{
			Try
				{	
					$env:Test_Admin_Account = "True"						
					get-aduser $Current_User -Credential $Creds										
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Admin_Window, "Success :-)", "Admin password has been setted successfully")	
				}
			Catch 
				{
					$env:Test_Admin_Account = "False"	
					remove-item $File -force			
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Admin_Window, "Oops :-(", "Admin password has not been setted successfully.`n`nCheck user name or password.")				
				}		
		}
		

})





############################################################################################################
# 								ACTION ON CLICK ON ABOUT
############################################################################################################

# ---------------------------------------------------------------------
# Action after clicking on About
# ---------------------------------------------------------------------

$Menu_About.Add_Click({
	$window_About.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$window_About.Width)
	$window_About.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$window_About.Height)
	$window_About.Show()
	$window_About.Activate()
	
	If ($Current_Version -lt $Tool_Gallery_Version)
		{
			$Main_Update_Status.Background = "Orange"
			$Update_Status_Label.Content = "A new version is available"
			$Update_Status_Label.Foreground = "White"
			$Update_Status_Label.FontWeight="Bold"
			$Update_Status_Label.Fontsize = "12"
			$Update_Status_Label.Margin = "0,7,0,0"		
		}
})





# Close the window if it's double clicked
# $Users_Window.Add_MouseDoubleClick({
	# $Users_Window.Hide()
# })

# Close the window if it loses focus
$Users_Window.Add_Deactivated({
	$Users_Window.Hide()	
	# Close_modal_progress	
})

$Computers_Window.Add_Deactivated({
	$Computers_Window.Hide()
})











# ----------------------------------------------------
# Part - About Window
# ----------------------------------------------------

# Close the window if it's double clicked
$window_About.Add_MouseDoubleClick({
	$window_About.Hide()
})

# Close the window if it loses focus
$window_About.Add_Deactivated({
	$window_About.Hide()
})


# ----------------------------------------------------
# Part - Exit
# ----------------------------------------------------

# When Exit is clicked, close everything and kill the PowerShell process
$Menu_Restart_Tool.add_Click({
	$Restart = "Yes"
	start-process -WindowStyle hidden powershell.exe "C:\ProgramData\Quick_AD_Support\QADS.ps1 '$Restart'" 	

	$Main_Tool_Icon.Visible = $false
	$window.Close()
	Stop-Process $pid	
 })
 

# ----------------------------------------------------
# Part - Exit
# ----------------------------------------------------

# When Exit is clicked, close everything and kill the PowerShell process
$Menu_Exit.add_Click({
	$Main_Tool_Icon.Visible = $false
	$window.Close()
	Stop-Process $pid
 })
 













# ----------------------------------------------------
# Part - Main action
# ----------------------------------------------------		


If (test-path "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\ActiveDirectory")
	{
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

		If ((Get-Module ActiveDirectory) -eq $null) 
		{
			Try
				{
					Import-Module ActiveDirectory -ErrorAction Stop -warningvariable ModuleError
					If ($ModuleError -ne $null)
						{
							[System.Windows.Forms.MessageBox]::Show("$ModuleError")									
							exit
						}
				}
			Catch
				{
					[System.Windows.Forms.MessageBox]::Show("Error while importig AD module")		
					exit
				}
		}
	}
Else
	{
		[System.Windows.Forms.MessageBox]::Show("ActiveDirectory module has not been found.")	
		exit
	}



################################################################################################################################"
# ACTIONS ON THE DIFFERENT WINDOW FROM CONTEXT MENU
################################################################################################################################"


$User_TxtBox.Add_MouseDoubleClick({
	$User_TxtBox.Text = ""
})


# $Expander_More.Add_PreviewMouseLeftButtonDown({	
	# $Expander_Basic.IsExpanded  = $false
	# $Expander_GroupMembership.IsExpanded  = $false
	# $Expander_Change_Password.IsExpanded  = $false
	# $Expander_Warning.IsExpanded  = $false
	# $Expander_More_options.IsExpanded  = $false
# })



# $Expander_Basic.Add_PreviewMouseLeftButtonDown({	
	# $Expander_More.IsExpanded  = $false
	# $Expander_GroupMembership.IsExpanded  = $false
	# $Expander_Change_Password.IsExpanded  = $false
	# $Expander_Warning.IsExpanded  = $false
	# $Expander_More_options.IsExpanded  = $false
# })


# $User_Display_Groups.Add_PreviewMouseLeftButtonDown({	
	# $Expander_More.IsExpanded  = $false
	# $Expander_Basic.IsExpanded  = $false
	# $Expander_Change_Password.IsExpanded  = $false
	# $Expander_Warning.IsExpanded  = $false
# })


# $Change_Password.Add_PreviewMouseLeftButtonDown({	
	# $Expander_More.IsExpanded  = $false
	# $Expander_GroupMembership.IsExpanded  = $false
	# $Expander_Basic.IsExpanded  = $false
	# $Expander_Warning.IsExpanded  = $false
# })


# $Expander_More_options.Add_PreviewMouseLeftButtonDown({	
	# $Expander_More.IsExpanded  = $false
	# $Expander_GroupMembership.IsExpanded  = $false
	# $Expander_Basic.IsExpanded  = $false
	# $Expander_Warning.IsExpanded  = $false
	# $Expander_Change_Password.IsExpanded  = $false
# })






############################################################################################################
# 											CHANGE PASSWORD PART
############################################################################################################

# ---------------------------------------------------------------------
# Action after clicking on change user password
# ---------------------------------------------------------------------

$Change_Password.Add_Click({	
	If ($Expander_GroupMembership.IsExpanded -eq $False)	
		{
			$Expander_Change_Password.Background = ""
		}
	Else
		{
			$Expander_Change_Password.Background = "Gray"		
		}

	If (($Expander_Change_Password.IsExpanded) -eq $True)
		{
			$Expander_Change_Password_Block.Visibility = "Collapsed"
			$Expander_Change_Password.IsExpanded = $False		
		}
	Else
		{
			$Expander_Change_Password_Block.Visibility = "Visible"
			$Expander_Change_Password.IsExpanded = $True		
		}
})




# ---------------------------------------------------------------------
# Action after clicking on set the user password
# ---------------------------------------------------------------------

$Set_Password.Add_Click({		
	$User = $Get_User_Infos.SamAccountName
	$First_PWD = $First_Password.Password.ToString()
	$Second_PWD = $Confirm_Password.Password.ToString()
	
	If (($First_Password.Password) -eq "")
		{
			$First_Password.Borderbrush = "Red"
		}
		
	If (($Confirm_Password.Password) -eq "")
		{
			$Confirm_Password.Borderbrush = "Red"
		}

	# $Use_Creds_Status = $false
	# $ProgData = $env:PROGRAMDATA
	# $QADS_Tool = "$ProgData\Quick_AD_Support\Infos"
	# $File = "$QADS_Tool\QADS_SD.txt"	
	# $File_User = "$QADS_Tool\QADS_User.txt"
	
	
	If ((test-path $File_User) -and (test-path $File)) 
		{
			$Admin_User_Name = Get-content ($File_User)	
			$Use_Creds_Status = $true
			[Byte[]] $Global:key = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
			$Global:Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, (Get-Content $File | ConvertTo-SecureString -Key $key)																		
		}	
		
	# If (test-path $File_User) 
		# {
			# $Admin_User_Name = Get-content ($File_User)	
			# $Use_Creds_Status = $true
			# $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, ($Encrypted_Password | ConvertTo-SecureString -Key $key)				
		# }			

	If ((($First_Password.Password) -ne "") -and (($Confirm_Password.Password) -ne ""))
		{
			If (($First_Password.Password) -eq ($Confirm_Password.Password))
				{
					$First_Password.Borderbrush = "Blue"
					$Confirm_Password.Borderbrush = "Blue"	

					$okAndCancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
					$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Users_Window,"Change password","Do you want really to change user password ?",$okAndCancel)
					If($result -eq "Affirmative")
						{
							Try	
								{
									If ($Use_Creds_Status -eq $true)
										{							
											$newpwd = ConvertTo-SecureString -String $First_PWD -AsPlainText -Force
                                            get-aduser $User | Set-ADAccountPassword -NewPassword $newpwd -Reset -Credential $Creds	
										}
									Else
										{
											$newpwd = ConvertTo-SecureString -String $First_PWD -AsPlainText Force
                                            get-aduser $User | Set-ADAccountPassword -NewPassword $newpwd -Reset 										
										}
										
									[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Success :-)", "Password has been successfully updated")
									User_Found_Custo									
								}
							Catch
								{							
									[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Oops :-(", "Can not change the password.`n`nCheck your rights or go to the admin part.")
								}									
						}					
				}
			Else
				{
					$First_Password.Borderbrush = "Red"
					$Confirm_Password.Borderbrush = "Red"				
				}
		}
	$Users_Window.Activate()
	$Users_Window.Show()			
})


############################################################################################################
# 											DISPLAY GROUPS PART
############################################################################################################

# ---------------------------------------------------------------------
# Action after clicking on display users group
# ---------------------------------------------------------------------

$User_Display_Groups.Add_Click({	
	If ($Expander_Change_Password.IsExpanded -eq $False)	
		{
			$Expander_GroupMembership.Background = ""
		}
	Else
		{
			$Expander_GroupMembership.Background = "Gray"		
		}

	If (($Expander_GroupMembership.IsExpanded) -eq $True)
		{
			$Expander_GroupMembership_Block.Visibility = "Collapsed"
			$Expander_GroupMembership.IsExpanded = $False		
		}
	Else
		{
			$Expander_GroupMembership_Block.Visibility = "Visible"
			$Expander_GroupMembership.IsExpanded = $True		
		}
})



# ---------------------------------------------------------------------
# Action after clicking on display computer group
# ---------------------------------------------------------------------

$Computer_Display_Groups.Add_Click({	
	$Computer_Expander_GroupMembership.Background = ""
	$Computer_Expander_GroupMembership.Background = "Gray"		

	If (($Computer_Expander_GroupMembership.IsExpanded) -eq $True)
		{
			$Computer_Expander_GroupMembership_Block.Visibility = "Collapsed"
			$Computer_Expander_GroupMembership.IsExpanded = $False		
		}
	Else
		{
			$Computer_Expander_GroupMembership_Block.Visibility = "Visible"
			$Computer_Expander_GroupMembership.IsExpanded = $True		
		}
})




############################################################################################################
# 											ENABLE / DISABLE ACCOUNT PART
############################################################################################################

# ---------------------------------------------------------------------
# Action after clicking on enable user account
# ---------------------------------------------------------------------

$User_Enable_Account.Add_Click({
	$User = $Get_User_Infos.SamAccountName
	$Account_Status = $Get_User_Infos.Enabled
		
	# $Use_Creds_Status = $false
	# $ProgData = $env:PROGRAMDATA
	# $QADS_Tool = "$ProgData\Quick_AD_Support\Infos"
	# $File = "$QADS_Tool\QADS_SD.txt"	
	# $File_User = "$QADS_Tool\QADS_User.txt"
	
	If ((test-path $File_User) -and (test-path $File)) 
		{
			$Admin_User_Name = Get-content ($File_User)	
			$Use_Creds_Status = $true
			[Byte[]] $Global:key = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
			$Global:Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, (Get-Content $File | ConvertTo-SecureString -Key $key)																		
		}	
		
	# If (test-path $File_User) 
		# {
			# $Admin_User_Name = Get-content ($File_User)	
			# $Use_Creds_Status = $true
			# $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, ($Encrypted_Password | ConvertTo-SecureString -Key $key)				
		# }			

	If (($Get_User_Infos.Enabled) -eq "True")
		{
			$okAndCancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
			$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Users_Window,"Disable account","Do you want to disable the account $User ?",$okAndCancel)
	
			If($result -eq "Affirmative")
				{
					Try	
						{		
							If ($Use_Creds_Status -eq $true)
								{
									Disable-ADAccount $User	-Credential $Creds		
								}
							Else
								{
									Disable-ADAccount $User	
								}
								
							[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Success :-)", "Account $User has been successfully disabled")	
					
						}
					Catch
						{							
							[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Oops :-(", "Can not disable the account.`n`nCheck your rights or go to the admin part.")							
						}
				}
			Else
				{
				}
			$Users_Window.Activate()
			$Users_Window.Show()					
		}
	Else
		{
			$okAndCancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
			$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Users_Window,"Enable account","Do you want to enable the account $User ?",$okAndCancel)
			If($result -eq "Affirmative")
				{
					Try	
						{
						
							If ($Use_Creds_Status -eq $true)
								{
									Enable-ADAccount $User	-Credential $Creds								
								}
							Else
								{
									Enable-ADAccount $User								
								}
								
							[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Success :-)", "Account $User has been successfully enabled")	
						}
					Catch
						{													
							[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Oops :-(", "Can not enable the account.`n`nCheck your rights or go to the admin part.")													
						}								
				}
			Else
				{
				}
			$Users_Window.Activate()
			$Users_Window.Show()					
		}				
})


# ---------------------------------------------------------------------
# Action after clicking on enable computer account
# ---------------------------------------------------------------------

$Computer_Enable_Account.Add_Click({
	$computer = $Get_Computer_Infos.Name
	$Account_Status = $Get_Computer_Infos.Enabled
	
	# $Use_Creds_Status = $false
	# $ProgData = $env:PROGRAMDATA
	# $QADS_Tool = "$ProgData\Quick_AD_Support\Infos"
	# $File = "$QADS_Tool\QADS_SD.txt"	
	# $File_User = "$QADS_Tool\QADS_User.txt"

	If ((test-path $File_User) -and (test-path $File)) 	
		{
			$Admin_User_Name = Get-content ($File_User)	
			$Use_Creds_Status = $true
			[Byte[]] $Global:key = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
			$Global:Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, (Get-Content $File | ConvertTo-SecureString -Key $key)																		
		}	
		
	# If (test-path $File_User) 
		# {
			# $Admin_User_Name = Get-content ($File_User)	
			# $Use_Creds_Status = $true
			# $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, ($Encrypted_Password | ConvertTo-SecureString -Key $key)				
		# }			

	If (($Get_Computer_Infos.Enabled) -eq "True")	
		{
			$okAndCancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
			$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Computers_Window,"Disable account","Do you want to disable the computer $computer ?",$okAndCancel)
			If($result -eq "Affirmative")
				{
					Try	
						{
						
							If ($Use_Creds_Status -eq $true)
								{
									Get-ADComputer $computer | disable-adaccount -Credential $Creds			
								}
							Else
								{
									Get-ADComputer $computer | disable-adaccount	
								}
								
							[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Success :-)", "Computer $computer has been successfully disabled")																											
						}
					Catch
						{
							[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Oops :-(", "Can not disable the account.`n`nCheck your rights or go to the admin part.")																			
						}
				}
			Else
				{
				}					
		}
	Else
		{
			$okAndCancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
			$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Computers_Window,"Enable account","Do you want to enable the computer $computer ?",$okAndCancel)
			If($result -eq "Affirmative")
				{
					Try	
						{
							If ($Use_Creds_Status -eq $true)
								{
									Get-ADComputer $computer | Enable-adaccount -Credential $Creds			
								}
							Else
								{
									Get-ADComputer $computer | Enable-adaccount								
								}
								
							[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Success :-)", "Computer $computer has been successfully enabled")																			
						}
					Catch
						{
							[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Oops :-(", "Can not enable the account.`n`nCheck your rights or go to the admin part.")																									
						}
				}
			Else
				{
				}
		}
	$Computers_Window.Activate()
	$Computers_Window.Show()			
})




############################################################################################################
# 											DELETE ACCOUNT PART
############################################################################################################

# ---------------------------------------------------------------------
# Action after clicking on delete user account
# ---------------------------------------------------------------------

$Delete_User_Account.Add_Click({
	$User = $Get_User_Infos.SamAccountName
	
	# $Use_Creds_Status = $false
	# $ProgData = $env:PROGRAMDATA
	# $QADS_Tool = "$ProgData\Quick_AD_Support\Infos"
	# $File = "$QADS_Tool\QADS_SD.txt"	
	# $File_User = "$QADS_Tool\QADS_User.txt"

	If ((test-path $File_User) -and (test-path $File)) 	
		{
			$Admin_User_Name = Get-content ($File_User)	
			$Use_Creds_Status = $true
			[Byte[]] $Global:key = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
			$Global:Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, (Get-Content $File | ConvertTo-SecureString -Key $key)																		
		}		
		
	# If (test-path $File_User) 
		# {
			# $Admin_User_Name = Get-content ($File_User)	
			# $Use_Creds_Status = $true
			# $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, ($Encrypted_Password | ConvertTo-SecureString -Key $key)				
		# }			
	
	$okAndCancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
	$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Users_Window,"Delete account","Do you want to delete the account $User ?",$okAndCancel)
	If($result -eq "Affirmative")
		{
			Try	
				{
					If ($Use_Creds_Status -eq $true)
						{
							Remove-ADUser $User	-confirm:$false	-Credential $Creds		
						}
					Else
						{
							Remove-ADUser $User	-confirm:$false		
						}
						
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Success :-)", "Account $User has been successfully deleted")	
					User_Found_Custo					
				}
			Catch
				{							
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Oops :-(", "Can not delete the account.`n`nCheck your rights or go to the admin part.")													
				}							
		}
	Else
		{
		}		
	$Users_Window.Activate()
	$Users_Window.Show()			
})


# ---------------------------------------------------------------------
# Action after clicking on delete computer account
# ---------------------------------------------------------------------

$Delete_Computer_Account.Add_Click({
	$Computer = $Get_Computer_Infos.Name
	
	# $Use_Creds_Status = $false
	# $ProgData = $env:PROGRAMDATA
	# $QADS_Tool = "$ProgData\Quick_AD_Support\Infos"
	# $File = "$QADS_Tool\QADS_SD.txt"	
	# $File_User = "$QADS_Tool\QADS_User.txt"
	
	If ((test-path $File_User) -and (test-path $File)) 
		{
			$Admin_User_Name = Get-content ($File_User)	
			$Use_Creds_Status = $true
			[Byte[]] $Global:key = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
			$Global:Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, (Get-Content $File | ConvertTo-SecureString -Key $key)																		
		}

	# If (test-path $File_User) 
		# {
			# $Admin_User_Name = Get-content ($File_User)	
			# $Use_Creds_Status = $true
			# $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, ($Encrypted_Password | ConvertTo-SecureString -Key $key)				
		# }			
	
	$okAndCancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
	$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Computers_Window,"Delete account","Do you want to delete the account $Computer ?",$okAndCancel)
	If($result -eq "Affirmative")
		{
			Try	
				{
					If ($Use_Creds_Status -eq $true)
						{
							Remove-ADComputer $Computer	-confirm:$false	-Credential $Creds		
						}
					Else
						{
							Remove-ADComputer $Computer	-confirm:$false								
						
						}

					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Success :-)", "Account $Computer has been successfully deleted")									
				}
			Catch
				{							
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Oops :-(", "Can not delete the account.`n`nCheck your rights or go to the admin part.")													
				}						
		}
	Else
		{
		}		
	$Computers_Window.Activate()
	$Computers_Window.Show()			
})




############################################################################################################
# 											UNLOCK ACCOUNT PART
############################################################################################################

# ---------------------------------------------------------------------
# Action after clicking on unlock user account
# ---------------------------------------------------------------------

$User_Unlock_Account.Add_Click({
	$User = $Get_User_Infos.SamAccountName
	$Users_Window.Activate()
	$Users_Window.Show()	
	$okAndCancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
	$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Users_Window,"Unlock account","Do you want to unlock the account $User ?",$okAndCancel)
	
	# $Use_Creds_Status = $false
	# $ProgData = $env:PROGRAMDATA
	# $QADS_Tool = "$ProgData\Quick_AD_Support\Infos"
	# $File = "$QADS_Tool\QADS_SD.txt"	
	# $File_User = "$QADS_Tool\QADS_User.txt"
	
	If ((test-path $File_User) -and (test-path $File)) 
		{
			$Admin_User_Name = Get-content ($File_User)	
			$Use_Creds_Status = $true
			[Byte[]] $Global:key = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
			$Global:Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, (Get-Content $File | ConvertTo-SecureString -Key $key)																		
		}			

	# If (test-path $File_User) 
		# {
			# $Admin_User_Name = Get-content ($File_User)	
			# $Use_Creds_Status = $true
			# $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, ($Encrypted_Password | ConvertTo-SecureString -Key $key)				
		# }			
	
	If($result -eq "Affirmative")
		{
			Try	
				{
					If ($Use_Creds_Status -eq $true)
						{
							Unlock-ADAccount -Identity $User -Credential $Creds	
						}
					Else
						{
							Unlock-ADAccount -Identity $User
						}	
						
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Oops :-(", "User $User has been successfully unlocked")																		
				}
			Catch
				{
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Oops :-(", "Can not unlock the account $User.`n`nCheck your rights or go to the admin part.")													
				}	
		}
	Else
		{
		}		
	$Users_Window.Activate()
	$Users_Window.Show()		
})


# ---------------------------------------------------------------------
# Action after clicking on unlock computer account
# ---------------------------------------------------------------------

$Computer_Unlock_Account.Add_Click({
	$computer = $Get_Computer_Infos.Name
	$Computers_Window.Activate()
	$Computers_Window.Show()	
	
	# $Use_Creds_Status = $false
	# $ProgData = $env:PROGRAMDATA
	# $QADS_Tool = "$ProgData\Quick_AD_Support\Infos"
	# $File = "$QADS_Tool\QADS_SD.txt"	
	# $File_User = "$QADS_Tool\QADS_User.txt"
	
	If ((test-path $File_User) -and (test-path $File)) 
		{
			$Admin_User_Name = Get-content ($File_User)	
			$Use_Creds_Status = $true
			[Byte[]] $Global:key = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
			$Global:Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, (Get-Content $File | ConvertTo-SecureString -Key $key)																		
		}			

	# If (test-path $File_User) 
		# {
			# $Admin_User_Name = Get-content ($File_User)	
			# $Use_Creds_Status = $true
			# $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Admin_User_Name, ($Encrypted_Password | ConvertTo-SecureString -Key $key)				
		# }			
	
	$okAndCancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
	$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Computers_Window,"Unlock account","Do you want to unlock the account $computer ?",$okAndCancel)
	If($result -eq "Affirmative")
		{
			Try	
				{			
					If ($Use_Creds_Status -eq $true)
						{
							Unlock-ADAccount -Identity $computer -Credential $Creds		
						}
					Else
						{
							Unlock-ADAccount -Identity $computer
						}
																										
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Oops :-(", "Computer $computer has been successfully unlocked")																		
				}
			Catch
				{
					[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Oops :-(", "Can not unlock the account $computer.`n`nCheck your rights or go to the admin part.")													
				}	
		}
	Else
		{
		}		
	$Computers_Window.Activate()
	$Computers_Window.Show()		
})





############################################################################################################
# 								CHECK USER / COMPUTER PROPERTIES PART
############################################################################################################

# ---------------------------------------------------------------------
# Action after clicking on check user properties
# ---------------------------------------------------------------------

$Check_User_btn.Add_Click({
	$Global:MyUser = $User_TxtBox.Text.ToString()
	$DatagGrid_Users.items.Clear()
	$DatagGrid_Groups.items.Clear()

	If ($MyUser -ne "")
		{			
			# Launch_modal_progress			
			Try
				{
					$MyUser = "*$MyUser*"
					$Global:Get_User_Infos = Get-ADUser -Filter {(name -like $MyUser) -or (SamAccountName -like $MyUser)} -properties *

					If ($Get_User_Infos -ne $null)
						{
							If ($Get_User_Infos.length -gt 1)
								{
									Warning_Users_Block	
									$Warning_Label.Content = "Several entries have been found for $MyUser"

									$DatagGrid_MultiUsers.Visibility = "Visible"
									ForEach ($User in $Get_User_Infos)				
										{
											$Users_values = New-Object PSObject
											$Users_values = $Users_values | Add-Member Account $User.SamAccountName -passthru		
											$Users_values = $Users_values | Add-Member Name $User.Name -passthru							
											$DatagGrid_Users.Items.Add($Users_values) > $null		
										}
										
									# Close_modal_progress	
									$Users_Window.Activate()
									$Users_Window.Show()									
								}
							Else
								{									
									User_Found_Custo									
									# Close_modal_progress	
									$Users_Window.Activate()
									$Users_Window.Show()
								}
									
						}
					Else
						{
							Clear_Users_Datas													
							Warning_Users_Block		
							$Warning_Label.Content = "Can not find the user $MyUser"		
							$DatagGrid_MultiUsers.Visibility = "Collapsed"
							
							# Close_modal_progress		
							$Users_Window.Activate()
							$Users_Window.Show()							
						}		
				}
			Catch
				{	
				
				
				
					Warning_Users_Block					
					$Warning_Label.Content = "Error with the command"														
					# Close_modal_progress			
					$Users_Window.Activate()
					$Users_Window.Show()					
				}				
		}
	Else
		{
			[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Oops :-(", "Please select a user !!!")				
		}		
})

# $Users_Window.Add_MouseLeftButtonDown({

# })	
	


# ---------------------------------------------------------------------
# Action after clicking on check computer properties
# ---------------------------------------------------------------------

$Check_Computer_btn.Add_Click({
	$MyComputer = $Computer_TxtBox.Text.ToString()
	$Computer_DatagGrid_Groups.items.Clear()

	If ($MyComputer -ne "")
		{		
			# Launch_modal_progress
			$MyComputer = "*$MyComputer*"
			$Global:Get_Computer_Infos = Get-ADComputer -Filter {(name -like $MyComputer)} -properties *	

				
			If ($Get_Computer_Infos -ne $null)
				{	
					If ($Get_Computer_Infos.length -gt 1)
						{
							Warning_Computers_Block								
							$Computer_Warning_Label.Content = "Several entries have been found for $MyUser"

							$DatagGrid_MultiComputers.Visibility = "Visible"
							ForEach ($Comp in $Get_Computer_Infos)				
								{
									$Computers_values = New-Object PSObject
									$Computers_values = $Computers_values | Add-Member Name $Comp.Name -passthru	
									$Computers_values = $Computers_values | Add-Member IP $Comp.IPv4Address -passthru	
									$Computers_values = $Computers_values | Add-Member OS $Comp.OperatingSystemVersion -passthru																									
									$DatagGrid_Computers.Items.Add($Computers_values) > $null		
								}
								
							# Close_modal_progress	
							$Computers_Window.Activate()
							$Computers_Window.Show()									
						}
					Else
						{									
							Computer_Found_Custo									
							# Close_modal_progress	
							$Computers_Window.Activate()
							$Computers_Window.Show()
						}							
				}
			Else
				{
					Clear_Users_Datas													
					Warning_Computers_Block		
					$Computer_Warning_Label.Content = "Can not find the computer $MyComputer"						
					# Close_modal_progress		
					$Computers_Window.Activate()
					$Computers_Window.Show()							
				}		
		
		}
	Else
		{
			[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Oops :-(", "Please select a user !!!")				
		}		

})




############################################################################################################
# 								CHECK USER / COMPUTER FROM DATAGRID PART
############################################################################################################

# ---------------------------------------------------------------------
# Action after clicking onthe user datagrid
# ---------------------------------------------------------------------

$DatagGrid_Users.Add_MouseDoubleClick({
	ForEach ($Selected_item in $DatagGrid_Users.Selecteditems)
		{	
			$LogonAccount = $Selected_item.Account
		}

	$Global:Get_User_Infos = Get-ADUser -Filter {(SamAccountName -like $LogonAccount)} -properties *
	User_Found_Custo											
})



# ---------------------------------------------------------------------
# Action after clicking onthe computer datagrid
# ---------------------------------------------------------------------

$DatagGrid_Computers.Add_MouseDoubleClick({
	ForEach ($Selected_item in $DatagGrid_Computers.Selecteditems)
		{	
			$Computer = $Selected_item.Name
		}

	$Global:Get_Computer_Infos = Get-ADComputer -Filter {(name -like $Computer)} -properties *	
	Computer_Found_Custo											
})




# ---------------------------------------------------------------------
# Action after clicking on the User GUI
# ---------------------------------------------------------------------

$Users_Window.Add_MouseDoubleClick({
})

$Users_Window.Add_MouseLeftButtonDown({
})





############################################################################################################
# 								EXPORT INFOS PART
############################################################################################################

# ---------------------------------------------------------------------
# Action after clicking on the export user informations
# ---------------------------------------------------------------------

$Export_User_Values.Add_Click({
	$tmp_folder = $env:TEMP	
	$Users_Infos_TXT = "$tmp_folder\QADS_Infos_$MyUser.txt"

	If ($User_TxtBox.Text -eq "")
		{
			[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Users_Window, "Oops :-(", "Please select a user !!!")						
		}
	Else
		{		
		
			$Global:Users_Infos_To_Export = @{
			'QADS_SamAccountName' = $Get_User_Infos.SamAccountName
			'QADS_FullName' = $Get_User_Infos.Name
			'QADS_UserOU' = $Get_User_Infos.DistinguishedName	
			'QADS_LastLogOn' = $Get_User_Infos.LastLogonDate	
			'QADS_IsLocked' = $Get_User_Infos.LockedOut					
			'QADS_Dept' = $Get_User_Infos.Department	 
			'QADS_IsEnabled' = $Get_User_Infos.Enabled			
			'QADS_LastBadPWD' = $Get_User_Infos.LastBadPasswordAttempt									
			'QADS_Mail' = $Get_User_Infos.Mail	
			'QADS_Office' = $Get_User_Infos.Office	
			'QADS_IsPWDExpired' = $Get_User_Infos.PasswordExpired									
			'QADS_PWDLastChange' = $Get_User_Infos.PasswordLastSet		
			'QADS_WhenChanged' = $Get_User_Infos.WhenChanged									
			'QADS_WhenCreated' = $Get_User_Infos.WhenCreated	
			'QADS_CannotChangePassword' = $Get_User_Infos.CannotChangePassword						
			}
			New-Object -Type PSObject -Property $Users_Infos_To_Export		
			$Users_Infos_To_Export | out-file $Users_Infos_TXT			
			invoke-item $Users_Infos_TXT		
		}
})

		
$Computer_Export_Values.Add_Click({
	$tmp_folder = $env:TEMP	
	$Computer_Infos_TXT = "$tmp_folder\QADS_Infos_$MyComputer.txt"

	If ($Computer_TxtBox.Text -eq "")
		{
			[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Computers_Window, "Oops :-(", "Please select a user !!!")						
		}
	Else
		{
			$Global:Computer_Infos_To_Export = @{
			'QADS_FullName' = $Get_Computer_Infos.Name
			'QADS_OU' = $Get_Computer_Infos.DistinguishedName	
			'QADS_LastLogOn' = $Get_Computer_Infos.LastLogonDate	
			'QADS_IsLocked' = $Get_Computer_Infos.LockedOut					
			'QADS_SID' = $Get_Computer_Infos.SID	 
			'QADS_IsEnabled' = $Get_Computer_Infos.Enabled			
			'QADS_OSVersion' = $Get_Computer_Infos.OperatingSystemVersion	
			'QADS_OS' = $Get_Computer_Infos.OperatingSystem									
			'QADS_ServicePack' = $Get_Computer_Infos.OperatingSystemServicePack									
			'QADS_Location' = $Get_Computer_Infos.Location	
			'QADS_Description' = $Get_Computer_Infos.Description	
			
			'QADS_IP' = $Get_Computer_Infos.IPv4Address	
			'QADS_DNSHostName' = $Get_Computer_Infos.DNSHostName	
			'QADS_IsPWDExpired' = $Get_Computer_Infos.PasswordExpired									
			'QADS_PWDLastChange' = $Get_Computer_Infos.PasswordLastSet		
			'QADS_WhenChanged' = $Get_Computer_Infos.WhenChanged									
			'QADS_WhenCreated' = $Get_Computer_Infos.WhenCreated	
			'QADS_ObjectCategory' = $Get_Computer_Infos.ObjectCategory			
			'QADS_CanonicalName' = $Get_Computer_Infos.CanonicalName			
			'QADS_PrimaryGroup' = $Get_Computer_Infos.PrimaryGroup			
			
			}
			New-Object -Type PSObject -Property $Computer_Infos_To_Export		
			$Computer_Infos_To_Export | out-file $Computer_Infos_TXT			
			invoke-item $Computer_Infos_TXT	
		}
})




# Make PowerShell Disappear
$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

# Force garbage collection just to start slightly lower RAM usage.
[System.GC]::Collect()



# Create an application context for it to all run within.
# This helps with responsiveness, especially when clicking Exit.
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)